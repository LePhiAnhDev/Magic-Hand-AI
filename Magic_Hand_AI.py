import os
import cv2
import mediapipe as mp
import numpy as np
import time
import threading
import queue
from collections import deque
import pyautogui
import traceback
import platform
import comtypes
import logging
import warnings
import sys

# Try to import pyfiglet and colorama for enhanced ASCII art banner
try:
    import pyfiglet
    from colorama import init, Fore, Back, Style
    init(autoreset=True)  # Initialize colorama
    pyfiglet_available = True
    colorama_available = True
except ImportError:
    pyfiglet_available = False
    colorama_available = False
    print("For a better experience, install pyfiglet and colorama: pip install pyfiglet colorama")
    
# Try to import termcolor for additional color options
try:
    from termcolor import colored
    termcolor_available = True
except ImportError:
    termcolor_available = False

# Suppress MediaPipe warnings
os.environ["MEDIAPIPE_DISABLE_GPU"] = "1"  # Force CPU to avoid some warnings
logging.getLogger("absl").setLevel(logging.ERROR)  # Suppress absl warnings
warnings.filterwarnings("ignore", category=UserWarning)  # Suppress general warnings

# Import and initialize system volume control library (Windows)
try:
    from ctypes import cast, POINTER
    from comtypes import CLSCTX_ALL
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
    volume_lib_available = True
except ImportError:
    volume_lib_available = False
    print("Could not import pycaw library. Volume will be controlled using shortcut keys.")
    print("To install: pip install pycaw comtypes")

# Try to import Selenium
try:
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from webdriver_manager.chrome import ChromeDriverManager
    selenium_available = True
except ImportError:
    selenium_available = False
    print("Could not import Selenium library. Please install: pip install selenium webdriver-manager")

# Setup MediaPipe Hands with optimized configuration
mp_hands = mp.solutions.hands
hands = mp_hands.Hands(
    max_num_hands=2,
    min_detection_confidence=0.7,  # Increase detection accuracy
    min_tracking_confidence=0.7,   # Increase tracking accuracy
    static_image_mode=False
)
mp_drawing = mp.solutions.drawing_utils
mp_drawing_styles = mp.solutions.drawing_styles

# Optimized queue with small size
frame_queue = queue.Queue(maxsize=1)
result_queue = queue.Queue(maxsize=1)

# Global variables
current_volume = 50                # Internal volume
system_volume = 50                 # Actual system volume
last_volume_change_time = 0
fps_values = deque(maxlen=10)      # Reduced size for faster response
processing_active = True

# Variables for YouTube playback speed control
current_speed = 1.0
target_speed = 1.0
last_speed_change_time = 0
prev_left_hand_distance = None
speed_values = [0.25, 0.5, 0.75, 1.0, 1.25, 1.5, 1.75, 2.0]
speed_index = 3  # Initial speed = 1.0
speed_direction_bias = 0  # To track change trend

# Distance history (for motion prediction)
distance_history = deque(maxlen=5)
filtered_distance_history = deque(maxlen=20)  # Store filtered values

# Global variables for Selenium driver
driver = None
video_url = None
selenium_active = False
browser_type = "chrome"  # Default browser type

# Initialize system volume control
volume_controller = None
if volume_lib_available:
    try:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume_controller = cast(interface, POINTER(IAudioEndpointVolume))
        
        # Read current volume
        volume_range = volume_controller.GetVolumeRange()
        min_vol, max_vol = volume_range[0], volume_range[1]
        vol = volume_controller.GetMasterVolumeLevelScalar()
        system_volume = int(vol * 100)
        current_volume = system_volume  # Update internal volume
        print(f"Connected to system volume control. Current volume: {system_volume}%")
    except Exception as e:
        print(f"Could not initialize volume control: {e}")
        volume_controller = None

# Advanced noise reduction filter for hand gestures
class AdvancedSmoothFilter:
    def __init__(self, alpha=0.5, responsiveness=0.5, min_alpha=0.2, max_alpha=0.95):
        self.value = None
        self.base_alpha = alpha
        self.responsiveness = responsiveness  # Sensitivity to changes (0-1)
        self.min_alpha = min_alpha  # Minimum alpha
        self.max_alpha = max_alpha  # Maximum alpha
        self.velocity = 0  # Rate of change
        self.acceleration = 0  # Acceleration of change
        self.last_values = deque(maxlen=3)  # Store recent values
    
    def update(self, new_value):
        if self.value is None:
            self.value = new_value
            self.last_values.append(new_value)
            return new_value
        
        # Calculate velocity and acceleration
        old_velocity = self.velocity
        self.velocity = new_value - self.value
        self.acceleration = self.velocity - old_velocity
        
        # Adjust alpha based on change magnitude and direction
        diff = abs(new_value - self.value)
        direction = 1 if new_value > self.value else -1
        
        # Decrease alpha for fast changes (for faster reaction)
        # Increase alpha for slow changes (for more stability)
        adjusted_alpha = max(self.min_alpha, 
                             min(self.max_alpha, 
                                 self.base_alpha - diff * self.responsiveness * direction))
        
        # Apply filter with adjusted alpha
        filtered_value = adjusted_alpha * new_value + (1 - adjusted_alpha) * self.value
        
        # Enhanced motion prediction for near-zero latency response
        prediction_factor = 0.4  # Slightly reduced prediction factor
        predicted_value = filtered_value + self.velocity * prediction_factor + self.acceleration * 0.15
        
        # Apply prediction with bounds checking to prevent overshooting
        max_deviation = 0.12  # Slightly reduced max deviation
        if abs(predicted_value - filtered_value) > max_deviation:
            # Limit prediction range but keep direction
            direction = 1 if predicted_value > filtered_value else -1
            predicted_value = filtered_value + (direction * max_deviation)
            
        # Return predicted value for speed control instead of filtered value
        # This creates a more immediate response
        if self.responsiveness > 0.8:  # Identify speed filter by high responsiveness
            self.value = filtered_value
            self.last_values.append(filtered_value)
            return predicted_value
        else:
            # For volume and other controls, use normal filtered value
            self.value = filtered_value
            self.last_values.append(filtered_value)
            return filtered_value

# Optimized filters for each gesture type
# For volume: More stable, less responsive
distance_filter = AdvancedSmoothFilter(alpha=0.7, responsiveness=0.3, min_alpha=0.3, max_alpha=0.9)
# For playback speed: Responsive but not too sensitive 
left_hand_filter = AdvancedSmoothFilter(alpha=0.2, responsiveness=0.85, min_alpha=0.05, max_alpha=0.5)  # Slightly reduced sensitivity

def display_fancy_banner():
    """Display a fancy colorful banner with LePhiAnhDev text"""
    
    if not (pyfiglet_available and (colorama_available or termcolor_available)):
        print("\n========== LePhiAnhDev ==========")
        print("Install pyfiglet, colorama, and termcolor for a fancy banner!")
        return None
    
    # List of impressive fonts to try in order of preference
    fancy_fonts = [
        'slant', 'banner3-D', 'isometric1', 'doom', 'larry3d', 'epic', 
        'colossal', 'graffiti', 'standard'
    ]
    
    # Try fonts until we find one that works
    banner_text = None
    used_font = None
    
    for font in fancy_fonts:
        try:
            banner_text = pyfiglet.figlet_format("LePhiAnhDev", font=font)
            used_font = font
            break
        except Exception:
            continue
    
    # If no fancy font worked, use default
    if banner_text is None:
        banner_text = pyfiglet.figlet_format("LePhiAnhDev")
    
    # Add colors using colorama if available
    if colorama_available:
        # Add gradient colors to each line
        colored_lines = []
        lines = banner_text.split('\n')
        
        # Generate gradient colors
        colors = [
            Fore.BLUE, Fore.CYAN, Fore.GREEN, Fore.YELLOW, 
            Fore.RED, Fore.MAGENTA, Fore.BLUE
        ]
        
        for i, line in enumerate(lines):
            if line.strip():
                # Create shadow effect by printing the line twice with offset
                shadow_line = line.replace('#', ' ').replace('/', ' ').replace('\\', ' ')
                
                # Use modulo to cycle through colors for gradient effect
                color_idx = i % len(colors)
                colored_line = colors[color_idx] + Style.BRIGHT + line
                colored_lines.append(colored_line)
            else:
                colored_lines.append(line)
        
        banner_text = '\n'.join(colored_lines)
    
    # Print decorative box around banner
    terminal_width = 80  # Default width
    try:
        # Try to get actual terminal width
        terminal_width = os.get_terminal_size().columns
    except:
        pass
    
    # Create decorative border
    border_top = "╔" + "═" * (terminal_width - 2) + "╗"
    border_bottom = "╚" + "═" * (terminal_width - 2) + "╝"
    
    # Print the banner with borders
    print("\n" + border_top)
    print(banner_text)
    
    # Print subtitle with different color if colorama is available
    if colorama_available:
        subtitle = "AI Hand Controller by LePhiAnhDev"
        padding = (terminal_width - len(subtitle) - 4) // 2
        print(Fore.CYAN + Style.BRIGHT + "║" + " " * padding + subtitle + " " * (terminal_width - len(subtitle) - padding - 4) + "║")
    
    print(border_bottom + "\n")
    
    return used_font

def get_browser_user_data_dir(browser_type="chrome"):
    """Return the default path to the browser's user data directory"""
    system = platform.system()
    
    if browser_type.lower() == "brave":
        if system == "Windows":
            return os.path.join(os.environ["LOCALAPPDATA"], "BraveSoftware", "Brave-Browser", "User Data")
        elif system == "Darwin":  # macOS
            return os.path.expanduser("~/Library/Application Support/BraveSoftware/Brave-Browser")
        elif system == "Linux":
            return os.path.expanduser("~/.config/BraveSoftware/Brave-Browser")
    else:  # Chrome (default)
        if system == "Windows":
            return os.path.join(os.environ["LOCALAPPDATA"], "Google", "Chrome", "User Data")
        elif system == "Darwin":  # macOS
            return os.path.expanduser("~/Library/Application Support/Google/Chrome")
        elif system == "Linux":
            return os.path.expanduser("~/.config/google-chrome")
    
    return None

def setup_selenium():
    """Initialize Chrome or Brave browser and open YouTube"""
    global driver, video_url, selenium_active, browser_type
    
    if not selenium_available:
        print("Selenium is not available - skipping browser initialization")
        return False
    
    print("\nInitializing browser to control YouTube...")
    print("\n*** IMPORTANT: Please ensure that your browser is closed before making a selection. ***\n")
    
    # Choose browser type
    browser_choice = input("Select browser (1 for Chrome, 2 for Brave, Enter for Chrome): ")
    browser_type = "brave" if browser_choice == "2" else "chrome"
    
    # Get path to browser's user data
    default_user_data_dir = get_browser_user_data_dir(browser_type)
    if default_user_data_dir and os.path.exists(default_user_data_dir):
        print(f"Found default User Data directory: {default_user_data_dir}")
    else:
        print(f"Default User Data directory for {browser_type.capitalize()} not found")
        default_user_data_dir = None
    
    # Allow user to input custom path
    custom_user_data_dir = input(f"Enter path to {browser_type.capitalize()} User Data directory (Press Enter to use {'default' if default_user_data_dir else 'no profile'}): ")
    
    # Use custom or default path
    user_data_dir = custom_user_data_dir if custom_user_data_dir else default_user_data_dir
    
    # Allow user to choose Profile
    if user_data_dir:
        profile = input("Enter Profile name (usually 'Default' or 'Profile 1', press Enter to use default): ")
    else:
        profile = None
    
    # Setup browser options
    if browser_type == "brave":
        options = webdriver.ChromeOptions()
        options.binary_location = ""  # Will be set based on OS
        
        # Detect Brave browser location based on OS
        if platform.system() == "Windows":
            brave_paths = [
                "C:\\Program Files\\BraveSoftware\\Brave-Browser\\Application\\brave.exe",
                "C:\\Program Files (x86)\\BraveSoftware\\Brave-Browser\\Application\\brave.exe",
                os.path.join(os.environ["LOCALAPPDATA"], "BraveSoftware", "Brave-Browser", "Application", "brave.exe")
            ]
            for path in brave_paths:
                if os.path.exists(path):
                    options.binary_location = path
                    break
        elif platform.system() == "Darwin":  # macOS
            options.binary_location = "/Applications/Brave Browser.app/Contents/MacOS/Brave Browser"
        elif platform.system() == "Linux":
            options.binary_location = "/usr/bin/brave-browser"
    else:
        options = webdriver.ChromeOptions()
    
    options.add_argument("--start-maximized")  # Maximize window for better visibility
    
    # Add option to use user data if available
    if user_data_dir:
        options.add_argument(f"--user-data-dir={user_data_dir}")
        if profile:
            options.add_argument(f"--profile-directory={profile}")
        print(f"Using User Data: {user_data_dir}")
        if profile:
            print(f"With Profile: {profile}")
    
    # Add option to keep browser open when selenium closes
    options.add_experimental_option("detach", True)
    
    # Options to reduce unnecessary notifications
    options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    options.add_experimental_option('useAutomationExtension', False)
    
    try:
        # Initialize driver
        if user_data_dir:
            # When using user-data-dir, shouldn't use webdriver_manager
            driver = webdriver.Chrome(options=options)
        else:
            # Use webdriver_manager when not using user-data-dir
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        
        # Open YouTube page
        video_url = input("\nEnter YouTube video URL (press Enter to use default video): ")
        if not video_url:
            video_url = "https://www.youtube.com/watch?v=dQw4w9WgXcQ"  # Default video
        
        driver.get(video_url)
        print(f"Opened YouTube video: {video_url}")
        
        # Wait for video to load
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.TAG_NAME, "video"))
        )
        
        # Automatically click play and skip ads if present
        try:
            time.sleep(3)  # Wait for ads if present
            
            # Try to find and click skip ad buttons
            try:
                skip_buttons = driver.find_elements(By.CSS_SELECTOR, ".ytp-ad-skip-button")
                if skip_buttons:
                    for button in skip_buttons:
                        driver.execute_script("arguments[0].click();", button)
                    print("Skipped ad")
            except:
                pass
            
            # Click on video to ensure it has focus
            video = driver.find_element(By.TAG_NAME, "video")
            driver.execute_script("arguments[0].click();", video)
            
            # Pause to ensure video has loaded
            time.sleep(0.5)
            
            # Click again to ensure video plays
            driver.execute_script("arguments[0].click();", video)
        except Exception as e:
            print(f"Warning when automatically playing video: {e}")
            print("Please click on the video in the browser to play")
        
        print("Successfully connected to YouTube!")
        
        # Add JavaScript to directly control
        inject_controller_script()
        selenium_active = True
        
        return True
    except Exception as e:
        print(f"Error initializing Selenium: {e}")
        traceback.print_exc()
        return False

def inject_controller_script():
    """Add JavaScript to YouTube page for ultra-fast playback speed control"""
    global driver
    
    if not driver:
        return False
    
    try:
        # Script optimized for high performance and ultra-low latency
        controller_script = """
        // Check if controller already exists
        if (!document.getElementById('ai-speed-controller')) {
            // Global variables
            window.aiHandController = {
                currentSpeed: document.querySelector('video').playbackRate,
                updateQueue: [],
                processingUpdate: false,
                lastUpdateTime: Date.now(),
                pendingAnimationFrame: null
            };
            
            // Create speed control panel
            const controlPanel = document.createElement('div');
            controlPanel.id = 'ai-speed-controller';
            controlPanel.style.position = 'fixed';
            controlPanel.style.bottom = '80px';
            controlPanel.style.right = '20px';
            controlPanel.style.backgroundColor = 'rgba(0, 0, 0, 0.7)';
            controlPanel.style.color = 'white';
            controlPanel.style.padding = '15px';
            controlPanel.style.borderRadius = '10px';
            controlPanel.style.zIndex = '9999';
            controlPanel.style.display = 'flex';
            controlPanel.style.flexDirection = 'column';
            controlPanel.style.alignItems = 'center';
            controlPanel.style.fontFamily = 'Arial, sans-serif';
            controlPanel.style.boxShadow = '0 4px 8px rgba(0,0,0,0.3)';
            controlPanel.style.transition = 'background-color 0.15s';
            
            // Title
            const title = document.createElement('div');
            title.textContent = 'AI Hand Controller';
            title.style.fontWeight = 'bold';
            title.style.fontSize = '14px';
            title.style.marginBottom = '10px';
            controlPanel.appendChild(title);
            
            // Display current speed
            const speedDisplay = document.createElement('div');
            speedDisplay.id = 'current-speed-display';
            speedDisplay.textContent = `Speed: ${window.aiHandController.currentSpeed.toFixed(2)}x`;
            speedDisplay.style.margin = '5px 0';
            speedDisplay.style.fontSize = '16px';
            controlPanel.appendChild(speedDisplay);
            
            // Control buttons
            const buttonContainer = document.createElement('div');
            buttonContainer.style.display = 'flex';
            buttonContainer.style.justifyContent = 'center';
            buttonContainer.style.width = '100%';
            buttonContainer.style.marginTop = '5px';
            
            const decreaseBtn = document.createElement('button');
            decreaseBtn.textContent = '-';
            decreaseBtn.style.margin = '0 5px';
            decreaseBtn.style.padding = '8px 15px';
            decreaseBtn.style.backgroundColor = '#c00';
            decreaseBtn.style.color = 'white';
            decreaseBtn.style.border = 'none';
            decreaseBtn.style.borderRadius = '5px';
            decreaseBtn.style.cursor = 'pointer';
            decreaseBtn.style.fontSize = '16px';
            decreaseBtn.style.fontWeight = 'bold';
            
            const increaseBtn = document.createElement('button');
            increaseBtn.textContent = '+';
            increaseBtn.style.margin = '0 5px';
            increaseBtn.style.padding = '8px 15px';
            increaseBtn.style.backgroundColor = '#c00';
            increaseBtn.style.color = 'white';
            increaseBtn.style.border = 'none';
            increaseBtn.style.borderRadius = '5px';
            increaseBtn.style.cursor = 'pointer';
            increaseBtn.style.fontSize = '16px';
            increaseBtn.style.fontWeight = 'bold';
            
            buttonContainer.appendChild(decreaseBtn);
            buttonContainer.appendChild(increaseBtn);
            controlPanel.appendChild(buttonContainer);
            
            // Add control panel to page
            document.body.appendChild(controlPanel);
            
            // Function to process update queue - reduce number of actual updates
            function processUpdateQueue() {
                if (window.aiHandController.processingUpdate) return;
                
                if (window.aiHandController.updateQueue.length > 0) {
                    window.aiHandController.processingUpdate = true;
                    
                    // Only get the latest item in queue
                    const latestRate = window.aiHandController.updateQueue.pop();
                    // Clear queue
                    window.aiHandController.updateQueue = [];
                    
                    const video = document.querySelector('video');
                    if (video) {
                        // Apply new speed immediately
                        video.playbackRate = latestRate;
                        window.aiHandController.currentSpeed = latestRate;
                        
                        // Update display in requestAnimationFrame for optimal performance
                        if (window.aiHandController.pendingAnimationFrame === null) {
                            window.aiHandController.pendingAnimationFrame = requestAnimationFrame(() => {
                                const display = document.getElementById('current-speed-display');
                                if (display) {
                                    display.textContent = `Speed: ${latestRate.toFixed(2)}x`;
                                }
                                // Flash effect
                                controlPanel.style.backgroundColor = 'rgba(204, 0, 0, 0.9)';
                                setTimeout(() => {
                                    controlPanel.style.backgroundColor = 'rgba(0, 0, 0, 0.7)';
                                }, 150);
                                
                                window.aiHandController.pendingAnimationFrame = null;
                            });
                        }
                    }
                    
                    window.aiHandController.processingUpdate = false;
                }
            }
            
            // Ultra-optimized update function
            window.updatePlaybackSpeed = function(rate) {
                // Add to queue
                window.aiHandController.updateQueue.push(rate);
                
                // Process immediately if possible
                processUpdateQueue();
                
                return true;
            };
            
            // Monitor playback speed changes from other sources (e.g. YouTube buttons)
            const video = document.querySelector('video');
            if (video) {
                video.addEventListener('ratechange', function() {
                    // Update if the change didn't come from us
                    if (Math.abs(video.playbackRate - window.aiHandController.currentSpeed) > 0.01) {
                        window.aiHandController.currentSpeed = video.playbackRate;
                        const display = document.getElementById('current-speed-display');
                        if (display) {
                            display.textContent = `Speed: ${video.playbackRate.toFixed(2)}x`;
                        }
                    }
                });
            }
            
            // Add events for buttons
            decreaseBtn.addEventListener('click', function() {
                const video = document.querySelector('video');
                if (video && video.playbackRate > 0.25) {
                    window.updatePlaybackSpeed(Math.max(0.25, video.playbackRate - 0.25));
                }
            });
            
            increaseBtn.addEventListener('click', function() {
                const video = document.querySelector('video');
                if (video && video.playbackRate < 2.0) {
                    window.updatePlaybackSpeed(Math.min(2.0, video.playbackRate + 0.25));
                }
            });
            
            // Current return function - ultra-optimized
            window.setYouTubeSpeed = function(speed) {
                if (speed !== window.aiHandController.currentSpeed) {
                    window.updatePlaybackSpeed(speed);
                }
                return true;
            };
            
            // Capture keyboard shortcuts
            document.addEventListener('keydown', function(e) {
                if (e.key === '.' || e.key === '>') {
                    const video = document.querySelector('video');
                    if (video && video.playbackRate < 2.0) {
                        window.updatePlaybackSpeed(Math.min(2.0, video.playbackRate + 0.25));
                    }
                } else if (e.key === ',' || e.key === '<') {
                    const video = document.querySelector('video');
                    if (video && video.playbackRate > 0.25) {
                        window.updatePlaybackSpeed(Math.max(0.25, video.playbackRate - 0.25));
                    }
                }
            });
            
            console.log('AI Hand Controller added to YouTube!');
        }
        """
        
        # Execute script
        driver.execute_script(controller_script)
        print("Added speed control panel to YouTube!")
        
        # Check default speed
        current_speed = driver.execute_script("return document.querySelector('video').playbackRate;")
        print(f"Current playback speed: {current_speed}x")
        
        return True
    except Exception as e:
        print(f"Error adding control panel: {e}")
        return False

def change_youtube_speed(new_speed):
    """Change YouTube playback speed with ultra-low latency"""
    global driver, selenium_active
    
    if not driver or not selenium_active:
        return False
    
    try:
        # Call the optimized JavaScript function
        driver.execute_script(f"return window.setYouTubeSpeed({new_speed});")
        return True
    except Exception as e:
        selenium_active = False  # Mark as no longer active
        return False

def predict_next_value(history, current_value, change_rate):
    """Predict next value based on history and change rate"""
    if len(history) < 2:
        return current_value
    
    # Calculate average change rate
    recent_changes = [history[i] - history[i-1] for i in range(1, len(history))]
    avg_change = sum(recent_changes) / len(recent_changes)
    
    # Predict next value
    predicted = current_value + avg_change * change_rate
    
    return predicted

def adjust_system_volume(target_volume_percent):
    """Adjust system volume directly (precise)"""
    global volume_controller, system_volume
    
    if volume_controller:
        try:
            # Limit range to 0-100%
            target_volume_percent = max(0, min(100, target_volume_percent))
            
            # Convert from percentage to 0-1 scale
            volume_scalar = target_volume_percent / 100.0
            
            # Set system volume directly
            volume_controller.SetMasterVolumeLevelScalar(volume_scalar, None)
            
            # Update tracking variable
            system_volume = target_volume_percent
            
            return system_volume
        except Exception as e:
            print(f"Error adjusting system volume: {e}")
    
    # Fallback method: use shortcut keys
    return adjust_volume_with_keys(target_volume_percent, system_volume)

def adjust_volume_with_keys(target, current):
    """Fallback method: adjust volume with shortcut keys"""
    diff = target - current
    if abs(diff) < 3:
        return current
    
    step_size = max(1, min(5, abs(diff) // 5))  # Dynamic step
    key = 'volumeup' if diff > 0 else 'volumedown'
    pyautogui.press(key, presses=step_size, interval=0.01)
    
    # Estimate new volume
    new_volume = current + (step_size * 2 if diff > 0 else -step_size * 2)
    new_volume = max(0, min(100, new_volume))
    
    return new_volume

def get_system_volume():
    """Read current system volume"""
    global volume_controller, system_volume
    
    if volume_controller:
        try:
            # Read volume from system
            vol = volume_controller.GetMasterVolumeLevelScalar()
            system_volume = int(vol * 100)
            return system_volume
        except Exception as e:
            print(f"Error reading system volume: {e}")
    
    # If can't read, return estimated value
    return system_volume

def adjust_playback_speed(direction, distance_change=None):
    """Adjust playback speed with ultra-low latency and prediction"""
    global speed_index, current_speed, speed_direction_bias, speed_values
    
    # Enhanced direct mapping approach
    if distance_change is not None:
        # Map distance change magnitude directly to speed change probability
        # This creates a much more continuous and natural control
        change_magnitude = abs(distance_change) * 100  # Scale up for better precision
        
        # Exponential mapping to make small movements more sensitive
        # but slightly reduced sensitivity
        change_probability = min(1.0, change_magnitude ** 1.4 / 35)  # Reduced sensitivity
        
        # Determine change direction and apply probability
        if distance_change > 0 and np.random.random() < change_probability and speed_index < len(speed_values) - 1:
            speed_index += 1
            current_speed = speed_values[speed_index]
            if selenium_active:
                change_youtube_speed(current_speed)
            return current_speed
        elif distance_change < 0 and np.random.random() < change_probability and speed_index > 0:
            speed_index -= 1
            current_speed = speed_values[speed_index]
            if selenium_active:
                change_youtube_speed(current_speed)
            return current_speed
    
    # Fallback to bias-based approach with slightly reduced sensitivity
    # Update current trend with high sensitivity
    if direction == "faster":
        speed_direction_bias += 1.8  # Slightly reduced responsiveness
    else:
        speed_direction_bias -= 1.8  # Slightly reduced responsiveness
    
    # Allow bias to build up faster but still have limits
    speed_direction_bias = max(-4.0, min(4.0, speed_direction_bias))
    
    # Use immediate response for deliberate hand movements
    should_change = False
    if direction == "faster" and speed_index < len(speed_values) - 1:
        if speed_direction_bias >= 1.2:  # Slightly higher threshold for better stability
            speed_index += 1
            speed_direction_bias = 0  # Reset after change
            should_change = True
    elif direction == "slower" and speed_index > 0:
        if speed_direction_bias <= -1.2:  # Slightly higher threshold for better stability
            speed_index -= 1
            speed_direction_bias = 0  # Reset after change
            should_change = True
    
    # If there's a change, update speed
    if should_change:
        current_speed = speed_values[speed_index]
        if selenium_active:
            change_youtube_speed(current_speed)
    
    return current_speed

def draw_centered_label(frame, text, position, size=0.5, thickness=1):
    """Draw centered label with white background and black text"""
    text_size = cv2.getTextSize(text, cv2.FONT_HERSHEY_SIMPLEX, size, thickness)[0]
    text_x, text_y = position
    
    # Draw white background
    bg_width = text_size[0] + 10
    bg_height = text_size[1] + 10
    bg_x = text_x - bg_width // 2
    bg_y = text_y - bg_height // 2
    
    cv2.rectangle(frame, (bg_x, bg_y), (bg_x + bg_width, bg_y + bg_height), (255, 255, 255), -1)
    
    # Draw black text centered on background
    text_offset_x = bg_x + (bg_width - text_size[0]) // 2
    text_offset_y = bg_y + (bg_height + text_size[1]) // 2
    cv2.putText(frame, text, (text_offset_x, text_offset_y), cv2.FONT_HERSHEY_SIMPLEX, size, (0, 0, 0), thickness)

def camera_reader():
    """Read camera from webcam (performance optimized)"""
    global processing_active
    
    try:
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            print("ERROR: Could not open webcam. Please check your camera connection.")
            processing_active = False
            return
            
        cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
        cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 360)
        cap.set(cv2.CAP_PROP_FPS, 60)  # Increase from 30fps to 60fps for better sensitivity
        cap.set(cv2.CAP_PROP_FOURCC, cv2.VideoWriter_fourcc(*'MJPG'))
        cap.set(cv2.CAP_PROP_BUFFERSIZE, 1)

        while processing_active:
            ret, frame = cap.read()
            if not ret:
                print("WARNING: Failed to capture frame from camera. Trying again...")
                time.sleep(0.1)
                continue
            
            frame = cv2.flip(frame, 1)
            try:
                if frame_queue.full():
                    frame_queue.get_nowait()
                frame_queue.put(frame, block=False)
            except queue.Full:
                pass
                
    except Exception as e:
        print(f"ERROR in camera thread: {e}")
        processing_active = False
    finally:
        if 'cap' in locals() and cap is not None:
            cap.release()
        print("Camera thread terminated.")

def hand_processor():
    """Process hand detection (optimized for performance and accuracy)"""
    global processing_active
    while processing_active:
        try:
            frame = frame_queue.get(timeout=0.03)  # Reduce wait time for faster response
            start_time = time.time()
            h, w, _ = frame.shape
            
            # Reduce processing size to increase performance
            scale = 0.5
            small_frame = cv2.resize(frame, (0, 0), fx=scale, fy=scale)
            rgb_frame = cv2.cvtColor(small_frame, cv2.COLOR_BGR2RGB)
            
            # Process hands
            results = hands.process(rgb_frame)
            
            processed_data = {
                'landmarks': [],
                'hand_sides': [],
                'hand_points': [],
                'left_hand_data': None,
                'frame': frame,
                'fps': 0
            }
            
            if results.multi_hand_landmarks and results.multi_handedness:
                # Loop through each detected hand
                for idx, (hand_landmarks, handedness) in enumerate(zip(results.multi_hand_landmarks, results.multi_handedness)):
                    # Extract hand information from recognition results
                    hand_side = 'left' if handedness.classification[0].label == "Left" else 'right'
                    processed_data['hand_sides'].append(hand_side)
                    processed_data['landmarks'].append(hand_landmarks)
                    
                    # Save coordinates of index finger and thumb
                    index_tip = hand_landmarks.landmark[8]  # Index finger
                    thumb_tip = hand_landmarks.landmark[4]  # Thumb
                    index_x, index_y = int(index_tip.x * w), int(index_tip.y * h)
                    thumb_x, thumb_y = int(thumb_tip.x * w), int(thumb_tip.y * h)
                    processed_data['hand_points'].append((index_x, index_y))
                    
                    # Save special information for left hand
                    if hand_side == 'left':
                        # Calculate distance between thumb and index finger on left hand
                        distance = np.hypot(thumb_x - index_x, thumb_y - index_y) / w
                        
                        # Save value to history (for prediction)
                        distance_history.append(distance)
                        
                        processed_data['left_hand_data'] = {
                            'index_point': (index_x, index_y),
                            'thumb_point': (thumb_x, thumb_y),
                            'distance': distance
                        }
            
            # Calculate FPS
            elapsed = max(time.time() - start_time, 0.001)
            fps_values.append(1.0 / elapsed)
            processed_data['fps'] = int(np.mean(fps_values))
            
            if result_queue.full():
                result_queue.get_nowait()
            result_queue.put(processed_data, block=False)
            
        except queue.Empty:
            time.sleep(0.001)  # Reduce wait time to increase response
        except Exception as e:
            print(f"Hand processor error: {e}")

def main():
    """Main program function"""
    global processing_active, current_volume, current_speed, prev_left_hand_distance
    global last_volume_change_time, last_speed_change_time, selenium_active, system_volume
    global filtered_distance_history
    
    # Display fancy banner
    used_font = display_fancy_banner()
    
    # Initialize variables to avoid errors
    current_speed = 1.0
    
    # Display welcome information
    print("===== AI HAND CONTROLLER WITH SELENIUM =====")
    print("This program will control YouTube playback speed directly")
    print("using hand gestures.")
    
    # Start processing threads (camera first to avoid delay)
    camera_thread = threading.Thread(target=camera_reader, daemon=True)
    processor_thread = threading.Thread(target=hand_processor, daemon=True)
    camera_thread.start()
    processor_thread.start()
    
    # Setup Selenium in a separate thread
    selenium_thread = threading.Thread(target=setup_selenium)
    selenium_thread.daemon = True
    selenium_thread.start()
    
    # Wait for Selenium to finish startup
    selenium_thread.join()
    
    print("\n===== USER GUIDE =====")
    print("1. Volume control: Use 2 hands (distance between two index fingers)")
    print("2. YouTube playback speed control:")
    print("   - Use left hand (distance between thumb-index finger)")
    print("   - Increase speed: Move thumb and index finger apart")
    print("   - Decrease speed: Pinch thumb and index finger together")
    print("\nSystem ready! Press ESC to exit.")
    
    # Create display window
    cv2.namedWindow('AI Hand Controller', cv2.WINDOW_NORMAL)
    
    # Update system volume
    system_volume = get_system_volume()
    current_volume = system_volume
    
    # Status tracking variables
    last_volume_status = ""
    last_speed_status = ""
    
    # Last system update time
    last_system_update = time.time()
    
    # Speed trend variable
    speed_trend = 0  # 0: no change, 1: increase, -1: decrease
    
    try:
        while True:
            try:
                # Check if processing is still active
                if not processing_active:
                    print("Processing has stopped. Exiting...")
                    break
                    
                try:
                    result = result_queue.get(timeout=0.01)  # Reduce timeout to increase response
                except queue.Empty:
                    # No frames available yet, check for exit key and continue
                    if cv2.waitKey(1) & 0xFF == 27:
                        break
                    continue
                
                frame = result['frame']
                hand_points = result['hand_points']
                landmarks = result['landmarks']
                hand_sides = result['hand_sides']
                left_hand_data = result['left_hand_data']
                fps = result['fps']
                
                h, w, _ = frame.shape
                
                # Re-read system volume every 1 second
                current_time = time.time()
                if current_time - last_system_update > 1.0:
                    system_volume = get_system_volume()
                    last_system_update = current_time
                
                # Draw landmarks
                for hand_landmark, hand_side in zip(landmarks, hand_sides):
                    # Display hand type with different colors
                    color = (0, 255, 0) if hand_side == 'left' else (0, 0, 255)
                    mp_drawing.draw_landmarks(
                        frame, 
                        hand_landmark, 
                        mp_hands.HAND_CONNECTIONS,
                        mp_drawing_styles.get_default_hand_landmarks_style(),
                        mp_drawing_styles.get_default_hand_connections_style()
                    )
                    
                    # Show hand label
                    wrist = hand_landmark.landmark[0]
                    wrist_x, wrist_y = int(wrist.x * w), int(wrist.y * h)
                    
                    # Use custom label function with white background
                    draw_centered_label(frame, f"{hand_side.capitalize()} hand", 
                                       (wrist_x, wrist_y - 15), size=0.5, thickness=1)
                
                # Process volume control (when 2 hands present)
                if len(hand_points) == 2:
                    x1, y1 = hand_points[0]
                    x2, y2 = hand_points[1]
                    distance = np.hypot(x2 - x1, y2 - y1) / w
                    smoothed_distance = distance_filter.update(distance)
                    
                    # Draw line between hands with more prominent visualization
                    cv2.line(frame, (x1, y1), (x2, y2), (255, 0, 0), 3)
                    
                    # Add mid-point volume display
                    mid_x = (x1 + x2) // 2
                    mid_y = (y1 + y2) // 2
                    
                    max_distance = 0.5
                    target_volume = int(np.interp(smoothed_distance, [0, max_distance], [0, 100]))
                    
                    # Draw centered volume display
                    draw_centered_label(frame, f"{system_volume}%", (mid_x, mid_y), size=0.6, thickness=2)
                    
                    # Volume bar directly on frame - centered vertically
                    bar_x = w - 50
                    bar_y = (h - 200) // 2  # Center vertically
                    bar_h = 200
                    bar_w = 30
                    
                    # Background
                    cv2.rectangle(frame, (bar_x, bar_y), (bar_x + bar_w, bar_y + bar_h), (200, 200, 200), -1)
                    
                    # Current volume
                    fill_h = int(bar_h * (system_volume / 100))
                    cv2.rectangle(frame, (bar_x, bar_y + bar_h - fill_h), (bar_x + bar_w, bar_y + bar_h),
                                 (0, 255, 0), -1)
                    
                    # Target volume indicator
                    target_y = bar_y + bar_h - int(bar_h * (target_volume / 100))
                    cv2.rectangle(frame, (bar_x - 5, target_y - 2), (bar_x + bar_w + 5, target_y + 2),
                                 (0, 0, 255), -1)
                    
                    # Volume percentage text
                    draw_centered_label(frame, f"{system_volume}%", (bar_x + bar_w // 2, bar_y + bar_h + 15), 0.5, 1)
                    
                    # Adjust system volume directly
                    current_time = time.time()
                    if current_time - last_volume_change_time > 0.1:  # 100ms
                        if abs(target_volume - system_volume) > 2:
                            old_volume = system_volume
                            system_volume = adjust_system_volume(target_volume)
                            last_volume_change_time = current_time
                            last_volume_status = "Increase" if system_volume > old_volume else "Decrease"
                    
                    # Show status
                    if last_volume_status:
                        color = (0, 255, 0) if last_volume_status == "Increase" else (0, 0, 255)
                        draw_centered_label(frame, last_volume_status, (bar_x + bar_w // 2, bar_y - 15), 0.5, 1)
                
                # Process playback speed control using left hand
                if left_hand_data:
                    index_point = left_hand_data['index_point']
                    thumb_point = left_hand_data['thumb_point']
                    distance = left_hand_data['distance']
                    
                    # Apply advanced smooth filter
                    smoothed_distance = left_hand_filter.update(distance)
                    filtered_distance_history.append(smoothed_distance)
                    
                    # Draw connection between index and thumb and highlight more
                    cv2.line(frame, index_point, thumb_point, (0, 255, 255), 3)
                    cv2.circle(frame, index_point, 10, (0, 255, 255), -1)
                    cv2.circle(frame, thumb_point, 10, (0, 255, 255), -1)
                    
                    # Display playback speed at midpoint
                    mid_x = (index_point[0] + thumb_point[0]) // 2
                    mid_y = (index_point[1] + thumb_point[1]) // 2
                    draw_centered_label(frame, f"{current_speed}x", (mid_x, mid_y), size=0.6, thickness=2)
                    
                    # Speed bar on left side - centered vertically
                    speed_bar_x = 50
                    speed_bar_y = (h - 200) // 2  # Center vertically
                    speed_bar_h = 200
                    speed_bar_w = 30
                    
                    # Background
                    cv2.rectangle(frame, (speed_bar_x, speed_bar_y), 
                                 (speed_bar_x + speed_bar_w, speed_bar_y + speed_bar_h), 
                                 (200, 200, 200), -1)
                    
                    # Current speed
                    normalized_speed = (speed_index / (len(speed_values) - 1))
                    fill_h = int(speed_bar_h * normalized_speed)
                    cv2.rectangle(frame, (speed_bar_x, speed_bar_y + speed_bar_h - fill_h), 
                                 (speed_bar_x + speed_bar_w, speed_bar_y + speed_bar_h),
                                 (255, 165, 0), -1)
                    
                    # Speed text
                    draw_centered_label(frame, f"{current_speed}x", 
                                      (speed_bar_x + speed_bar_w // 2, speed_bar_y + speed_bar_h + 15), 0.5, 1)
                    
                    if prev_left_hand_distance is not None:
                        distance_change = smoothed_distance - prev_left_hand_distance
                        
                        # Update general speed trend (for display)
                        if abs(distance_change) > 0.005:  # More sensitive to small changes
                            speed_trend = 1 if distance_change > 0 else -1
                        else:
                            speed_trend = 0
                        
                        # Near-zero threshold but slightly increased for stability
                        dynamic_threshold = 0.0025 + 0.002 * (1 - abs(distance_change) * 12)
                        dynamic_threshold = max(0.002, min(0.005, dynamic_threshold))  # Slightly increased threshold
                        
                        # Process speed change with a small delay for stability
                        current_time = time.time()
                        # Slightly longer delay (15ms) for better stability
                        if current_time - last_speed_change_time > 0.015:
                            if abs(distance_change) > dynamic_threshold:
                                direction = "faster" if distance_change > 0 else "slower"
                                
                                # Save previous state
                                old_speed = current_speed
                                
                                # Apply speed control with direct distance change input for more precision
                                current_speed = adjust_playback_speed(direction, distance_change)
                                
                                # If speed changes, update status
                                if current_speed != old_speed:
                                    last_speed_status = "Speed up" if current_speed > old_speed else "Slow down"
                                
                                # Update time to avoid continuous changes
                                last_speed_change_time = current_time
                        
                        # Show trend indicator near speed bar
                        trend_text = ""
                        if speed_trend > 0:
                            trend_text = "▲"
                            draw_centered_label(frame, trend_text, 
                                             (speed_bar_x + speed_bar_w // 2, speed_bar_y - 15), 0.7, 2)
                        elif speed_trend < 0:
                            trend_text = "▼"
                            draw_centered_label(frame, trend_text, 
                                             (speed_bar_x + speed_bar_w // 2, speed_bar_y - 15), 0.7, 2)
                    
                    # Update previous distance
                    prev_left_hand_distance = smoothed_distance
                    
                    # Show speed status
                    if last_speed_status:
                        color = (255, 165, 0) if last_speed_status == "Speed up" else (0, 165, 255)
                        draw_centered_label(frame, last_speed_status, 
                                         (speed_bar_x + speed_bar_w // 2, speed_bar_y - 35), 0.5, 1)
                
                # Display FPS
                cv2.putText(frame, f"FPS: {fps}", (w - 80, 20),
                           cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 1)
                
                # Display browser type
                browser_info = f"Using {browser_type.capitalize()}"
                cv2.putText(frame, browser_info, (10, 20),
                           cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 255), 1)
                
                # Show Selenium status
                status_text = "Connected" if selenium_active else "Disconnected"
                status_color = (0, 255, 0) if selenium_active else (0, 0, 255)
                cv2.putText(frame, f"YouTube: {status_text}", (10, h - 10),
                           cv2.FONT_HERSHEY_SIMPLEX, 0.5, status_color, 1)
                
                cv2.imshow('AI Hand Controller', frame)
                
                if cv2.waitKey(1) & 0xFF == 27:  # Exit with ESC
                    break
                    
            except queue.Empty:
                if cv2.waitKey(1) & 0xFF == 27:
                    break
            except Exception as e:
                print(f"Error in main loop: {e}")
                traceback.print_exc()
    
    except KeyboardInterrupt:
        print("\nProgram interrupted by user.")
    except Exception as e:
        print(f"\nUnexpected error occurred: {e}")
        traceback.print_exc()
    finally:
        # Clean up resources
        print("\nCleaning up resources...")
        processing_active = False
        
        try:
            # Ensure threads are properly terminated
            if 'camera_thread' in locals() and camera_thread.is_alive():
                camera_thread.join(timeout=1.0)
                
            if 'processor_thread' in locals() and processor_thread.is_alive():
                processor_thread.join(timeout=1.0)
                
            # Properly close OpenCV windows
            cv2.destroyAllWindows()
            
            # Correctly release MediaPipe resources
            if 'hands' in globals():
                hands.close()
                
            # Clean up volume controller
            if volume_controller is not None:
                pass  # Nothing special needed for pycaw cleanup
                
            print("\nClosed AI Hand Controller. Browser still open so you can continue watching video.")
        except Exception as cleanup_error:
            print(f"Error during cleanup: {cleanup_error}")
            
if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Fatal error: {e}")
        traceback.print_exc()
        print("\nProgram terminated due to an error. Please check your setup and try again.")
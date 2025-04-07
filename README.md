## Introduction
AI Hand Controller uses Computer Vision to recognize hand gestures and control various functions on your computer. The application can control system volume and YouTube video playback speed through hand gestures via webcam.

## Requirements
- Python 3.10.11 or higher
- Device with Webcam support, if you don't have one you can use Iriun Webcam to use your phone as a Webcam.
- Run the command to install the necessary libraries: ```pip install opencv-python mediapipe numpy pyautogui pyfiglet colorama termcolor pycaw comtypes selenium webdriver-manager```

- Run the command to install browser drivers: ```playwright install```

## Usage Guide
1. **Launch the application**: ```python Magic_Hand_AI.py```
2. **Volume control**:
   - Bring both hands into the frame, with index fingers pointing out
   - **Increase distance between hands** → **Increase volume**
   - **Decrease distance between hands** → **Decrease volume**
3. **Browser playback speed control**:
   - Bring your left hand into the frame, with thumb and index finger extended
   - **Increase distance between thumb and index finger** → **Increase playback speed**
   - **Decrease distance between thumb and index finger** → **Decrease playback speed**
4. **Exit application**: Press **ESC** key

## Features
- Control system volume by the distance between both hands
- Control browser playback speed by the distance between thumb and index finger of the left hand
- Direct integration with browsers (Chrome or Brave)
- Visual display with volume and speed bars

## Author
**Lê Phi Anh**  

## Contact for Work
- Discord: LePhiAnhDev  
- Telegram: @lephianh386ht  
- GitHub: [LePhiAnhDev](https://github.com/LePhiAnhDev)

from pynput import keyboard as kb
from pywinauto.application import Application
from time import sleep
import json
import os
import ctypes

user32 = ctypes.windll.user32

# Get the current directory path
dir_path = os.path.realpath(os.path.dirname(__file__))

# Read the configuration file
with open(os.path.join(dir_path, 'config.json')) as config_file:
    data = json.load(config_file)

# Initialize variables from the configuration file. Check the json file for more details.
active_key = data['activeKey'] # Default: =
skip_key = data['skipKey']  # Default: space
dialog_key = data['dialogKey']  # Default choice: 1
exit_key = data['exitKey'] # Default: [
pause_key = data['pauseKey'] # Default: ;
delay_press_time = data['delayPressTime'] # Default: 0.3
try:
    delay_press_time = float(delay_press_time)
except:
    print("Please use a number to define the delay time between press")
# Flags for pausing and exiting
pause_flag = True
exit_flag = False



def flush_input():
    """Clear the input buffer."""
    try:
        import msvcrt
        while msvcrt.kbhit():
            msvcrt.getch()
    except ImportError:
        import sys, termios  # For Linux/Unix
        termios.tcflush(sys.stdin, termios.TCIOFLUSH)


class DoQuest:
    """Class to perform quests."""

    @staticmethod
    def press_key(key):
        """Handle key press events."""
        global pause_flag, exit_flag
        if key == kb.KeyCode(char=pause_key):
            print("Script is paused.")
            pause_flag = True
        elif key == kb.KeyCode(char=active_key):
            print("Script is running.")
            pause_flag = False
        elif key == kb.KeyCode(char=exit_key):
            print("Exit")
            exit_flag = True
            flush_input()  # Clear the input buffer
            return False
        

    @staticmethod
    def stop_listener():
        """Stop listening for key presses."""
        global exit_flag
        exit_flag = True
        flush_input()  # Clear the input buffer
        return False

    @staticmethod
    def listener():
        """Start listening for key presses."""
        listener = kb.Listener(on_press=DoQuest.press_key)
        listener.start()

    @staticmethod
    def act(window):
        """Perform actions."""
        if not pause_flag:
            window.send_keystrokes(skip_key)
            sleep(delay_press_time)
            window.send_keystrokes(dialog_key)
            sleep(delay_press_time)
        else:
            sleep(1)


class Main:
    """Main class to run the program."""

    def __init__(self):
        try: 
            window = find_window()
        except Exception as error:
            print(error)
            self.exit_program()
        if window:
            self.run_auto(window)
        else:
            self.exit_program()

    @staticmethod
    def run_auto(window):
        """Run the auto."""
        print(f"\nPress key '{exit_key}' to stop the program.")
        print(f"Press key '{pause_key}' to Pause/Resume.")
        print(f"Press key '{active_key}' to enable the auto.")
        print("Note: These keys can be configured in config.json.")
        if skip_key == ' ': 
            _skip_key = 'space' 
        else: 
            _skip_key = skip_key
        #skip_key = 'space' if skip_key == ' ' else skip_key
        print(f"\nThe auto will keep sending {_skip_key} and {dialog_key} continuously when enabled.\n")
        DoQuest.listener()
        while not exit_flag:
            DoQuest.act(window)
            sleep(0.1)

    @staticmethod
    def exit_program():
        """Exit the program."""
        print("\nCannot find 'Honkai: Star Rail'. Are you sure the game is running?")
        for i in range(3, 0, -1):
            print(str(i) + " ")
            sleep(1)
        print("Exit")
        exit()

def set_window_text(hwnd, text):
    try:
        # Encode string to unicode
        text = ctypes.create_unicode_buffer(text)
        
        # Set title
        msg = user32.SetWindowTextW(hwnd,text)
    except Exception as error:
        print(f"Error: {error}")
        exit()

def find_window():
    """Find the application window."""
    current_window_name = input("\nEnter the current window name (or just enter if no changes, default = 'Honkai: Star Rail'): ").strip() or "Honkai: Star Rail"
    app = Application().connect(title=current_window_name, found_index=0)
    window = app.window(found_index=0, title_re=current_window_name)
    real_handle = window.handle
    new_window_name = input("If you want to change the title name of the current window, please enter the new window name (or just enter if no changes, default = 'Honkai: Star Rail'): ").strip() or "Honkai: Star Rail"
    set_window_text(real_handle, new_window_name)
    print(f"\nThe window name has been changed to ##### {new_window_name} #####")
    window = app.window(handle=real_handle, found_index=0)
    return window

if __name__ == "__main__":
    Main()

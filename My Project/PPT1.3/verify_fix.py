import win32com.client
import os
import time

def verify_ppt_env():
    print("Verifying PowerPoint environment...")
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = True
        print("Success: PowerPoint launched.")
        app.Quit()
        return True
    except Exception as e:
        print(f"Error launching PowerPoint: {e}")
        return False

if __name__ == "__main__":
    if verify_ppt_env():
        print("Verification PASSED: PowerPoint environment handles basic launch/quit.")
    else:
        print("Verification FAILED.")

import keyboard, subprocess
import pandas as pd
import pyautogui
from time import sleep
from os import environ, system
import schedule


def click_button(image_name, confidence=0.9):
    # Resolve button names to image paths in buttons folder
    image_path = f"buttons\\{image_name}.png"
    # Find, then click button
    location = pyautogui.locateCenterOnScreen(image_path, confidence=confidence)
    if location is None:
        # If button isn't immediately found, wait another ten seconds and try once more
        sleep(10)
        location = pyautogui.locateCenterOnScreen(image_path, confidence=confidence)
    assert location is not None, f"Could not find button: {image_path}"
    pyautogui.moveTo(location)
    pyautogui.click()


class Zoom:
    def join(self, id, password=""):
        print(f"Joining meeting {id}…")
        # Open the Zoom app
        subprocess.Popen(f"{environ['programfiles']}\\Zoom\\bin\\Zoom.exe")
        sleep(3)
        # Click the join button on the Zoom home screen
        click_button("home_join")
        sleep(1)
        # Type the meeting ID from the dataframe
        keyboard.write(id)
        sleep(1)

        # Click the join button on the join dialog
        click_button("dialog_join")
        sleep(7)

        # Enter Password if given
        if password:
            keyboard.write(password)
            sleep(3)
            # Click the join button on the password dialog
            click_button("password_dialog_join")

    def leave(self):
        print("Killing Zoom…")
        system("TASKKILL /F /IM Zoom.exe")

    def read_schedule(self, filename):
        weekdays = {
            "Mo": "monday",
            "Tu": "tuesday",
            "We": "wednesday",
            "Th": "thursday",
            "Fr": "friday",
            "Sa": "saturday",
            "Su": "sunday",
        }
        # Read CSV as strings only (no NaN on empty field)
        df = pd.read_csv(filename, dtype=str, keep_default_na=False)

        # Register scheduled join and leave tasks for every meeting
        for i, row in df.iterrows():
            getattr(schedule.every(), weekdays[row["Weekday"]]).at(
                row["Starting Time"]
            ).do(
                lambda id=row["Meeting ID"], password=row["Passcode"]: self.join(
                    id, password
                )
            )
            getattr(schedule.every(), weekdays[row["Weekday"]]).at(row["End Time"]).do(
                self.leave
            )


if __name__ == "__main__":
    Zoom().read_schedule("meetingschedule.csv")
    n = schedule.idle_seconds()
    while n is not None:
        if n >= 0:
            print(f"Going to sleep for {n} seconds")
            sleep(n)
            try:
                schedule.run_pending()
            except AssertionError as e:  # If a button can't be found, print error and carry on
                print(e)
        else:  # If a scheduled task is still running, n becomes negative. Wait it out
            sleep(30)
        n = schedule.idle_seconds()

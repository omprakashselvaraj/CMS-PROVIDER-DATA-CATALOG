import schedule
import time
import subprocess

def run_main_file():
    subprocess.call(["python", "main.py"])  # Replace "python" with your Python interpreter if needed

# Schedule the job
schedule.every().tuesday.at("20:27").do(run_main_file)

# The scheduler will continuously check if the job needs to be run
while True:
    schedule.run_pending()
    time.sleep(60)
import subprocess
from Config import *
# Replace 'your_executable.exe' with the path to your executable file
#executable_path = "D:/text document/LeadPush/LeadPush/bin/Debug/LeadPush.exe"
from dotenv import load_dotenv
import os

# Load environment variables from .env file
load_dotenv()

executable_path = os.getenv("Executable_path")
def run():
    try:
        # Run the executable
        subprocess.run(executable_path, shell=True, check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error running the executable: {e}")
        logging.error(f"Error running the executable: {e}")
    except FileNotFoundError:
        print(f"Executable not found at {executable_path}")
        logging.error(f"Executable not found at {executable_path}")



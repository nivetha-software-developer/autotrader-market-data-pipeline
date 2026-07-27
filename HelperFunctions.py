import json
import os
import sys
from datetime import datetime, timedelta
from pathlib import Path

from dotenv import load_dotenv


# Get absolute path of a file or dir
def get_abs_path(path: str) -> str:
    project_root = os.path.abspath(os.path.join(os.path.dirname(__file__)))
    return os.path.join(project_root, path)


root_folder = get_abs_path('')
print(f'Root folder: {root_folder}')

today = datetime.now()
is_linux = sys.platform.startswith('linux')
if is_linux:
    today = today + timedelta(hours=5, minutes=30)
    
todayDate = today.strftime("%Y%m%d")
todayDateInDash = today.strftime("%Y-%m-%d")
yesterday = today - timedelta(days=1)
yesterdayDate = yesterday.strftime("%Y%m%d")
yesterdayDateInDash = yesterday.strftime("%Y-%m-%d")

current_time = datetime.now().strftime("%H%M%S")

# Convert string to datetime object
yesterday_dt = datetime.strptime(yesterdayDate, "%Y%m%d")

# Subtract one day to get the day before yesterday
prev_date_dt = yesterday_dt - timedelta(days=1)

# Format back to string
prevDate = prev_date_dt.strftime("%Y%m%d")

# Set up logging
log_folder_path = os.path.join(root_folder, 'logs')
os.makedirs(log_folder_path, exist_ok=True)


def isProduction():
    return getattr(sys, 'frozen', False)


# set up Appsettings.json
appsettings_file_path = os.path.join(root_folder, 'Appsettings.json')
appsettings = json.loads(open(appsettings_file_path).read())


def get_appsettings(key):
    return appsettings.get(key)


FORCE_PRODUCTION = get_appsettings('FORCE_PRODUCTION')
IS_PRODUCTION = FORCE_PRODUCTION or isProduction()


def get_recent_folder(path):
    return max(os.listdir(os.path.join(root_folder, path)),
               key=lambda x: os.path.getctime(os.path.join(root_folder, path, x)))


env_path = Path(__file__).resolve().parent / ".env"
load_dotenv(dotenv_path=env_path)


def get_environment_variable(variable_name, default_value=None):
    variable_value = os.getenv(variable_name) or default_value
    return variable_value


import os
from pathlib import Path
from dotenv import load_dotenv

# 自动加载项目根目录下的 .env 文件
BASE_DIR = Path(__file__).resolve().parents[1]
dotenv_path = BASE_DIR / '.env'
if dotenv_path.exists():
    load_dotenv(dotenv_path)
else:
    print("⚠️ 未找到 .env 文件，将使用默认配置")

def get_env(key: str, default=None) -> str:
    """获取环境变量"""
    return os.getenv(key, default)

# 基础路径
BASE_DIR = Path(get_env("BASE_DIR", BASE_DIR))
CONFIG_DIR = BASE_DIR / get_env("CONFIG_DIR", "config")
LOG_DIR = BASE_DIR / get_env("LOG_DIR", "logs")
PERSISTENCE_DIR = BASE_DIR / get_env("PERSISTENCE_DIR", "persistence")
EXCEL_DIR = BASE_DIR / get_env("EXCEL_DIR", "data")

# 配置文件路径
CONFIG_PATH = CONFIG_DIR / get_env("CONFIG_FILE", "config.json")
PROGRESS_PATH = CONFIG_DIR / get_env("PROGRESS_FILE", "progress.json")

# 日志文件路径
LOG_FILE = LOG_DIR / "automation.log"

# 飞书数据
APP_ID = get_env("APP_ID")
APP_SECRET = get_env("APP_SECRET")
REDIRECT_URI = get_env("REDIRECT_URI")
SPREADSHEET_TOKEN = get_env("SPREADSHEET_TOKEN")
SHEET_ID = get_env("SHEET_ID")

#持久化数据
USER_DATA_FILE = get_env("USER_DATA_FILE")
TOKEN_STORE_FILE = get_env("TOKEN_STORE_FILE")
SHEET_STORE_FILE = get_env("SHEET_STORE_FILE")

for path in [LOG_DIR, CONFIG_DIR, PERSISTENCE_DIR, EXCEL_DIR]:
    path.mkdir(parents=True, exist_ok=True)

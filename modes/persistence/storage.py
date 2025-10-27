# storage.py
import json
from threading import RLock
from core.env import PERSISTENCE_DIR

class Storage:
    """
    简单 JSON 持久化存储模块
    - 每个实例对应一个 JSON 文件
    - 支持 get/set/delete 操作
    - 线程安全
    """
    def __init__(self, filename: str):
        self.filepath = PERSISTENCE_DIR / filename
        self.filepath.parent.mkdir(parents=True, exist_ok=True)
        self._lock = RLock()
        self._data = self._load_file()

    def _load_file(self):
        if self.filepath.exists():
            try:
                with open(self.filepath, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return {}
        else:
            return {}

    def _save_file(self):
        with open(self.filepath, "w", encoding="utf-8") as f:
            json.dump(self._data, f, ensure_ascii=False, indent=2)

    def get(self, key, default=None):
        with self._lock:
            return self._data.get(key, default)

    def set(self, key, value):
        with self._lock:
            self._data[key] = value
            self._save_file()

    def delete(self, key):
        with self._lock:
            if key in self._data:
                del self._data[key]
                self._save_file()

    def clear(self):
        with self._lock:
            self._data = {}
            self._save_file()

    def all(self):
        with self._lock:
            return dict(self._data)

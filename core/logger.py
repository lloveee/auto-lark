# core/logger.py
"""
日志管理模块
提供日志记录和回调功能
"""
from datetime import datetime
from typing import Callable, Optional
import threading
import logging


logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)


class Logger:
    """日志管理器"""

    def __init__(self):
        self._callbacks = []
        self._logs = []
        self._lock = threading.Lock()
        self._max_logs = 1000  # 最大日志条数

    def add_callback(self, callback: Callable[[str, str], None]):
        """添加日志回调函数"""
        with self._lock:
            self._callbacks.append(callback)

    def remove_callback(self, callback: Callable[[str, str], None]):
        """移除日志回调函数"""
        with self._lock:
            if callback in self._callbacks:
                self._callbacks.remove(callback)

    def _notify(self, level: str, message: str):
        """通知所有回调"""
        with self._lock:
            for callback in self._callbacks:
                try:
                    callback(level, message)
                except Exception as e:
                    print(f"日志回调错误: {e}")

    def _add_log(self, level: str, message: str):
        """添加日志到历史"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] [{level}] {message}"

        with self._lock:
            self._logs.append(log_entry)
            if len(self._logs) > self._max_logs:
                self._logs.pop(0)

    def info(self, message: str):
        """记录信息日志"""
        self._add_log("INFO", message)
        self._notify("INFO", message)

    def warning(self, message: str):
        """记录警告日志"""
        self._add_log("WARNING", message)
        self._notify("WARNING", message)

    def error(self, message: str):
        """记录错误日志"""
        self._add_log("ERROR", message)
        self._notify("ERROR", message)

    def success(self, message: str):
        """记录成功日志"""
        self._add_log("SUCCESS", message)
        self._notify("SUCCESS", message)

    def get_logs(self):
        """获取所有日志"""
        with self._lock:
            return self._logs.copy()

    def clear(self):
        """清空日志"""
        with self._lock:
            self._logs.clear()
        self.info("日志已清空")


# 全局日志实例
logger = Logger()
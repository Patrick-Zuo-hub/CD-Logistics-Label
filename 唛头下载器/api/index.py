# Vercel Python serverless 入口：复用上级目录的 Flask app
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app  # noqa: E402

# Vercel 的 Python runtime 会自动识别名为 app 的 WSGI 对象

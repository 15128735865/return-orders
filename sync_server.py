"""
本地同步服务器
- /ping          健康检查
- /sync          触发图片同步（SSE 流式进度）
- /sync-excel    手动触发 Excel 导入（SSE 流式进度）
- /excel-config  获取/设置 Excel 监控路径
- /status        服务器状态
- 自动监控 Excel 文件变化，变化时自动导入
"""
import http.server
import socketserver
import json
import sys
import os
import subprocess
import threading
import re
import time
import urllib.request
from urllib.parse import urlparse, parse_qs

HOST = 'localhost'
PORT = 8765
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, 'sync_config.json')

SUPABASE_URL = "https://jlzsonjjfgojmwgghxbl.supabase.co"
SERVICE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpsenNvbmpqZmdvam13Z2doeGJsIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc3NjkxNjY4OSwiZXhwIjoyMDkyNDkyNjg5fQ.qTAqHLBeWyVUfV8uxdP2-55EFI7kyh4aJ2RGJHrhQTo"
OZON_CLIENT_ID = "3306389"
OZON_API_KEY = "bd757233-6d64-4e7c-9d36-3db020b88533"


# ========== 配置管理 ==========
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {'excel_files': []}


def save_config(config):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


# ========== 文件监控 ==========
SYNC_LOG = []  # 最近 50 条同步记录
SYNC_LOG_MAX = 50

def add_sync_log(entry):
    entry['time'] = time.strftime('%H:%M:%S')
    SYNC_LOG.insert(0, entry)
    if len(SYNC_LOG) > SYNC_LOG_MAX:
        SYNC_LOG.pop()


class FileWatcher(threading.Thread):
    """轮询监控 Excel 文件变化，检测到变化自动导入"""

    def __init__(self, config_path=CONFIG_FILE):
        super().__init__(daemon=True)
        self.config_path = config_path
        self.mtimes = {}  # path -> last mtime
        self.running = True
        self.last_import_time = {}  # path -> last import time (防抖)
        self.debounce_seconds = 5
        self._log_file = os.path.join(SCRIPT_DIR, 'sync_events.log')

    def _log(self, msg):
        try:
            with open(self._log_file, 'a', encoding='utf-8') as f:
                f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n")
        except:
            pass
        print(msg)

    def run(self):
        self._log("[监控] 文件监控已启动")
        while self.running:
            try:
                config = load_config()
                for ef in config.get('excel_files', []):
                    if not ef.get('enabled', True):
                        continue
                    path = ef['path']
                    if not os.path.exists(path):
                        continue
                    mtime = os.path.getmtime(path)
                    if path in self.mtimes and mtime > self.mtimes[path]:
                        # 防抖：10 秒内不重复导入
                        last_import = self.last_import_time.get(path, 0)
                        if time.time() - last_import > self.debounce_seconds:
                            fname = os.path.basename(path)
                            self._log(f"[监控] 检测到文件变化: {fname}")
                            self.last_import_time[path] = time.time()
                            add_sync_log({'source': 'auto', 'file': fname, 'status': 'running'})
                            # 在后台线程运行导入
                            threading.Thread(target=self.import_excel, args=(path, ef.get('type', '')), daemon=True).start()
                    self.mtimes[path] = mtime
            except Exception as e:
                self._log(f"[监控] 错误: {e}")
            time.sleep(3)  # 每 3 秒检查一次

    def import_excel(self, filepath, excel_type):
        """后台导入 Excel"""
        fname = os.path.basename(filepath)
        try:
            env = os.environ.copy()
            env.update({
                'SUPABASE_URL': SUPABASE_URL,
                'SUPABASE_SERVICE_KEY': SERVICE_KEY,
                'EXCEL_PATH': filepath,
                'EXCEL_TYPE': excel_type,
                'PYTHONUNBUFFERED': '1'
            })
            proc = subprocess.run(
                [sys.executable, os.path.join(SCRIPT_DIR, 'sync_excel.py')],
                env=env,
                cwd=SCRIPT_DIR,
                capture_output=True,
                text=True,
                timeout=120
            )
            if proc.returncode == 0:
                last_line = ''
                lines = proc.stdout.strip().split('\n')
                if lines:
                    last_line = lines[-1]
                self._log(f"[监控] 导入完成: {fname} - {last_line}")
                # 解析结果
                import re
                new_match = re.search(r'新增\s*(\d+)\s*条', last_line)
                update_match = re.search(r'更新\s*(\d+)\s*条', last_line)
                add_sync_log({
                    'source': 'auto', 'file': fname, 'status': 'ok',
                    'new': int(new_match.group(1)) if new_match else 0,
                    'updated': int(update_match.group(1)) if update_match else 0
                })
            else:
                self._log(f"[监控] 导入失败: {fname} - {proc.stderr[:200]}")
                add_sync_log({'source': 'auto', 'file': fname, 'status': 'error', 'error': proc.stderr[:200]})
        except Exception as e:
            self._log(f"[监控] 导入异常: {fname} - {e}")
            add_sync_log({'source': 'auto', 'file': fname, 'status': 'error', 'error': str(e)})

    def stop(self):
        self.running = False


# ========== 核心：导入逻辑（供 sync_excel.py 和服务器共用）==========
def build_excel_import_env():
    return {
        'SUPABASE_URL': SUPABASE_URL,
        'SUPABASE_SERVICE_KEY': SERVICE_KEY,
        'PYTHONUNBUFFERED': '1'
    }


# ========== HTTP 服务器 ==========
class ThreadingHTTPServer(socketserver.ThreadingMixIn, http.server.HTTPServer):
    daemon_threads = True


class CORSHandler(http.server.BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass  # 减少控制台输出

    def send_cors(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def send_json(self, data, code=200):
        self.send_response(code)
        self.send_cors()
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode())

    def send_sse(self):
        self.send_response(200)
        self.send_cors()
        self.send_header('Content-Type', 'text/event-stream')
        self.send_header('Cache-Control', 'no-cache')
        self.send_header('Connection', 'keep-alive')
        self.end_headers()

    def sse_write(self, data):
        self.wfile.write(f"data: {json.dumps(data, ensure_ascii=False)}\n\n".encode())
        self.wfile.flush()

    def stream_subprocess(self, script_name, env_overrides=None):
        """运行 Python 脚本并流式输出 SSE 进度"""
        env = os.environ.copy()
        env.update(build_excel_import_env())
        if env_overrides:
            env.update(env_overrides)
        env['PYTHONUNBUFFERED'] = '1'

        try:
            proc = subprocess.Popen(
                [sys.executable, os.path.join(SCRIPT_DIR, script_name)],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                env=env,
                cwd=SCRIPT_DIR
            )

            for line in iter(proc.stdout.readline, ''):
                line = line.strip()
                if not line:
                    continue
                data = {'text': line}

                # 图片同步进度格式
                m = re.search(r'(\d+)/(\d+)\s*\((\d+)%\).*?累计\s*(\d+)\s*张', line)
                if m:
                    data.update({
                        'current': int(m.group(1)),
                        'total': int(m.group(2)),
                        'percent': int(m.group(3)),
                        'found': int(m.group(4)),
                        'type': 'progress'
                    })
                elif '进度' in line or '导入' in line or '条' in line:
                    # Excel 导入进度
                    data['type'] = 'progress'
                    pm = re.search(r'(\d+)/(\d+)', line)
                    if pm:
                        data['current'] = int(pm.group(1))
                        data['total'] = int(pm.group(2))
                        data['percent'] = int(pm.group(1)) * 100 // int(pm.group(2))

                if '完成' in line:
                    data['type'] = 'done'
                if 'error' in line.lower() or '失败' in line or 'SKIP' in line:
                    data['type'] = 'warn'

                self.sse_write(data)

            proc.wait()

            final = {'type': 'done', 'text': '操作完成'}
            self.sse_write(final)

        except Exception as e:
            self.sse_write({'type': 'error', 'text': str(e)})

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_cors()
        self.end_headers()

    def do_GET(self):
        path = urlparse(self.path).path

        if path == '/ping':
            return self.send_json({'status': 'ok', 'uptime': 'running'})

        if path == '/status':
            config = load_config()
            watching = []
            for ef in config.get('excel_files', []):
                exists = os.path.exists(ef['path'])
                mtime = os.path.getmtime(ef['path']) if exists else None
                watching.append({
                    'path': ef['path'],
                    'filename': os.path.basename(ef['path']),
                    'exists': exists,
                    'mtime': time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(mtime)) if mtime else None,
                    'enabled': ef.get('enabled', True)
                })
            return self.send_json({'status': 'ok', 'watching': watching})

        if path == '/sync-log':
            return self.send_json({'log': SYNC_LOG})

        if path == '/sync':
            # 图片同步 SSE
            self.send_sse()
            return self.stream_subprocess('sync_images.py', {
                'OZON_CLIENT_ID': OZON_CLIENT_ID,
                'OZON_API_KEY': OZON_API_KEY
            })

        if path == '/sync-excel':
            # Excel 手动导入 SSE
            self.send_sse()
            # 支持 ?path=xxx 指定单个文件
            qs = parse_qs(urlparse(self.path).query)
            filepath = qs.get('path', [None])[0]

            if filepath:
                # 指定了单个文件
                add_sync_log({'source': 'manual', 'file': os.path.basename(filepath), 'status': 'running'})
                env_overrides = {'EXCEL_PATH': filepath}
                return self.stream_subprocess('sync_excel.py', env_overrides)
            else:
                # 同步所有已配置的文件
                config = load_config()
                files = [ef['path'] for ef in config.get('excel_files', []) if ef.get('enabled', True) and os.path.exists(ef['path'])]
                if not files:
                    self.sse_write({'type': 'done', 'text': '没有可用的 Excel 文件，请先在设置中添加'})
                    return

                add_sync_log({'source': 'manual', 'file': f'{len(files)} 个文件', 'status': 'running'})
                self.sse_write({'type': 'progress', 'text': f'开始同步 {len(files)} 个文件...', 'current': 0, 'total': len(files), 'percent': 0})
                for idx, fp in enumerate(files):
                    fname = os.path.basename(fp)
                    self.sse_write({'type': 'progress', 'text': f'[{idx+1}/{len(files)}] 正在导入: {fname}', 'current': idx + 1, 'total': len(files), 'percent': (idx + 1) * 100 // len(files)})
                    env_overrides = {'EXCEL_PATH': fp}
                    self.stream_subprocess('sync_excel.py', env_overrides)

                # 刷新图片缓存
                self.sse_write({'type': 'progress', 'text': 'Excel 导入完成，开始检查新图片...', 'current': len(files), 'total': len(files), 'percent': 100})
                self.stream_subprocess('sync_images.py', {
                    'OZON_CLIENT_ID': OZON_CLIENT_ID,
                    'OZON_API_KEY': OZON_API_KEY
                })
                self.sse_write({'type': 'done', 'text': '全部完成！'})
                return

        self.send_json({'error': 'Not Found'}, 404)

    def do_POST(self):
        path = urlparse(self.path).path
        content_len = int(self.headers.get('Content-Length', 0))
        raw_body = self.rfile.read(content_len) if content_len > 0 else b'{}'

        # ===== 文件上传 =====
        if path == '/upload-excel':
            import base64
            try:
                body = json.loads(raw_body)
            except:
                return self.send_json({'error': 'Invalid request'}, 400)

            file_data = base64.b64decode(body.get('file_data', ''))
            target = body.get('target', '')  # 配置中的索引或路径
            filename = body.get('filename', '')
            config = load_config()

            # 确定保存路径
            save_path = None
            if target and target in ('0', '1', '2', '3', '4', '5'):
                idx = int(target)
                files = config.get('excel_files', [])
                if idx < len(files):
                    save_path = files[idx]['path']
            if not save_path:
                for ef in config.get('excel_files', []):
                    if ef['path'] == target:
                        save_path = target
                        break
            if not save_path and filename:
                save_path = os.path.join(os.path.expanduser('~'), 'Desktop', filename)

            if not save_path:
                return self.send_json({'error': '未找到对应的文件路径，请先在设置中配置'}, 400)

            os.makedirs(os.path.dirname(save_path), exist_ok=True)
            with open(save_path, 'wb') as f:
                f.write(file_data)

            # 确保路径在配置中
            found = False
            for ef in config.get('excel_files', []):
                if ef['path'] == save_path:
                    found = True
                    break
            if not found:
                config['excel_files'].append({'path': save_path, 'type': '', 'enabled': True})
                save_config(config)
            watcher.mtimes[save_path] = os.path.getmtime(save_path)

            # 立即触发导入
            self.send_json({'ok': True, 'filename': os.path.basename(save_path), 'path': save_path, 'imported': True})
            # 异步导入
            threading.Thread(target=watcher.import_excel, args=(save_path, ''), daemon=True).start()
            return

        # ===== JSON API =====
        try:
            body = json.loads(raw_body)
        except json.JSONDecodeError:
            return self.send_json({'error': 'Invalid JSON body'}, 400)

        if path == '/excel-config':
            action = body.get('action', 'get')

            if action == 'get':
                config = load_config()
                for ef in config.get('excel_files', []):
                    ef['filename'] = os.path.basename(ef['path'])
                    ef['exists'] = os.path.exists(ef['path'])
                    if ef['exists']:
                        ef['mtime'] = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(os.path.getmtime(ef['path'])))
                    else:
                        ef['mtime'] = None
                return self.send_json(config)

            if action == 'save':
                config = load_config()
                config['excel_files'] = body.get('excel_files', [])
                save_config(config)
                # 重置监控的文件时间戳，避免立即触发导入
                for ef in config['excel_files']:
                    if os.path.exists(ef['path']):
                        watcher.mtimes[ef['path']] = os.path.getmtime(ef['path'])
                return self.send_json({'ok': True, 'files': len(config['excel_files'])})

            if action == 'add':
                config = load_config()
                new_path = body.get('path', '').strip()
                # 去掉首尾引号（中英文引号 + 空格）
                new_path = new_path.strip('\u201c\u201d\u0022\u0027 \t\n')
                excel_type = body.get('type', '')
                # 统一斜杠方向并规范化
                new_path = os.path.normpath(new_path.replace('/', '\\'))
                # 写入调试日志
                with open(os.path.join(SCRIPT_DIR, 'server_debug.log'), 'a', encoding='utf-8') as f:
                    f.write(f'ADD PATH: {repr(new_path)} exists={os.path.exists(new_path)}\n')
                if not new_path or not os.path.exists(new_path):
                    return self.send_json({'error': f'文件不存在: {repr(new_path)}'}, 400)
                # 去重
                for ef in config['excel_files']:
                    if ef['path'] == new_path:
                        return self.send_json({'ok': True, 'message': '已存在'})
                config['excel_files'].append({
                    'path': new_path,
                    'type': excel_type,
                    'enabled': True
                })
                save_config(config)
                watcher.mtimes[new_path] = os.path.getmtime(new_path)
                return self.send_json({'ok': True, 'filename': os.path.basename(new_path)})

            if action == 'remove':
                config = load_config()
                remove_path = body.get('path', '')
                config['excel_files'] = [ef for ef in config['excel_files'] if ef['path'] != remove_path]
                save_config(config)
                return self.send_json({'ok': True})

            if action == 'toggle':
                config = load_config()
                toggle_path = body.get('path', '')
                for ef in config['excel_files']:
                    if ef['path'] == toggle_path:
                        ef['enabled'] = not ef.get('enabled', True)
                        break
                save_config(config)
                return self.send_json({'ok': True})

            return self.send_json({'error': '未知操作'}, 400)

        self.send_json({'error': 'Not Found'}, 404)


# ========== 启动 ==========
watcher = None


def main():
    global watcher

    # 启动文件监控
    watcher = FileWatcher()
    watcher.start()

    server = ThreadingHTTPServer((HOST, PORT), CORSHandler)
    print(f'同步服务器已启动: http://{HOST}:{PORT}')
    print(f'  /ping          - 健康检查')
    print(f'  /status        - 监控状态')
    print(f'  /sync          - 图片同步')
    print(f'  /sync-excel    - Excel 导入')
    print(f'  /excel-config  - Excel 配置')
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\n已停止')
        watcher.stop()
        server.shutdown()


if __name__ == '__main__':
    main()

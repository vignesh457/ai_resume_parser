import json
import io
from http.server import BaseHTTPRequestHandler
from urllib.parse import parse_qs, urlparse
import os
from upstash_redis import Redis

redis = Redis(url=os.environ.get("UPSTASH_REDIS_REST_URL"), token=os.environ.get("UPSTASH_REDIS_REST_TOKEN"))

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        query_components = parse_qs(urlparse(self.path).query)
        task_id = query_components.get("task_id", [None])[0]
        
        if not task_id:
            self.send_response(400)
            self.end_headers()
            return

        status = redis.get(f"status:{task_id}")
        
        if status == "COMPLETED":
            file_hex = redis.get(f"file:{task_id}")
            file_bytes = bytes.fromhex(file_hex)
            
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            self.send_header('Content-Disposition', f'attachment; filename="Tailored_Resume_{task_id[:8]}.docx"')
            self.end_headers()
            self.wfile.write(file_bytes)
        else:
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({"status": status or "NOT_FOUND"}).encode('utf-8'))
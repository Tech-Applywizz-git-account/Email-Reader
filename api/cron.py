from http.server import BaseHTTPRequestHandler
import sys
import os

# Ensure the parent directory is in the Python path so we can import email_to_supabase
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from email_to_supabase import check_inbox, logger

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        logger.info("=" * 60)
        logger.info("🚀 VERCEL CRON TRIGGERED: Email → Supabase Monitor")
        logger.info("=" * 60)
        
        try:
            check_inbox()
            
            self.send_response(200)
            self.send_header('Content-type', 'text/plain')
            self.end_headers()
            self.wfile.write("Cron execution successful".encode('utf-8'))
        except Exception as e:
            logger.error(f"💥 CRON ERROR: {str(e)}")
            self.send_response(500)
            self.send_header('Content-type', 'text/plain')
            self.end_headers()
            self.wfile.write(f"Cron execution failed: {str(e)}".encode('utf-8'))

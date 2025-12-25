from flask import Flask, render_template, request, jsonify, send_file
from app import run_email_extraction_workflow
import threading
import io
import sys
import os
from datetime import datetime

app = Flask(__name__)

# Global state for logs and execution
execution_in_progress = False
execution_result = None
execution_logs = []
log_stream = None

class LogCapture(io.StringIO):
    """Capture logs and print them"""
    def write(self, message):
        super().write(message)
        if message.strip():
            timestamp = datetime.now().strftime("%H:%M:%S")
            log_entry = f"[{timestamp}] {message.strip()}"
            execution_logs.append(log_entry)
            print(message, end='')  # Also print to console
        return len(message)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/execute', methods=['POST'])
def execute_task():
    """Execute the email extraction workflow"""
    global execution_in_progress, execution_result, execution_logs, log_stream
    
    if execution_in_progress:
        return jsonify({
            "success": False,
            "message": "Task already in progress"
        }), 400
    
    data = request.json
    gmail_account = data.get('gmail_account')
    gmail_password = data.get('gmail_password')
    
    # Validate inputs
    if not all([gmail_account, gmail_password]):
        return jsonify({
            "success": False,
            "message": "Gmail account and password are required"
        }), 400
    
    # Reset logs and state
    execution_in_progress = True
    execution_result = None
    execution_logs = []
    
    # Run task in background thread
    thread = threading.Thread(
        target=run_task_background,
        args=(gmail_account, gmail_password)
    )
    thread.daemon = True
    thread.start()
    
    return jsonify({
        "success": True,
        "message": "Task started"
    })

def run_task_background(gmail_account: str, gmail_password: str):
    """Run the extraction in background"""
    global execution_in_progress, execution_result
    
    try:
        # Just run the workflow directly without stdout redirection
        execution_result = run_email_extraction_workflow(gmail_account, gmail_password)
        execution_logs.append(f"[INFO] Workflow completed successfully")
        execution_logs.append(f"[INFO] Email count: {execution_result.get('email_count', 0)}")
        execution_logs.append(f"[INFO] File saved: {execution_result.get('output_path', 'N/A')}")
    except Exception as e:
        error_msg = str(e)
        execution_logs.append(f"[ERROR] Workflow failed: {error_msg}")
        execution_result = {
            "success": False,
            "error": error_msg,
            "message": f"Error: {error_msg}",
            "output_path": "",
            "email_count": 0
        }
    finally:
        execution_in_progress = False

@app.route('/api/status', methods=['GET'])
def get_status():
    """Get current execution status and logs"""
    return jsonify({
        "in_progress": execution_in_progress,
        "result": execution_result,
        "logs": execution_logs
    })

@app.route('/api/download', methods=['GET'])
def download_file():
    """Download the Excel file"""
    if execution_result and execution_result.get('output_path'):
        output_path = execution_result['output_path']
        if os.path.exists(output_path):
            return send_file(output_path, as_attachment=True, download_name='sent_emails.xlsx')
    return jsonify({"error": "File not found"}), 404

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=False)
#  gmail_account="dev.m.3771@gmail.com",
#         gmail_password="jmlr bxym ivud cjam"
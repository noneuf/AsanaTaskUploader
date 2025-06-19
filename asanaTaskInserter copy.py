import pandas as pd
import datetime
from flask import Flask, request, render_template_string, Response, send_file, stream_with_context
import os
from dotenv import load_dotenv
import requests
import traceback
import time

app = Flask(__name__)
load_dotenv()

HTML_TEMPLATE = '''
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Asana Task Scheduler</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
  <style>
    body { background-color: #f4f7fa; padding-top: 50px; }
    .container {
      max-width: 700px;
      background-color: #fff;
      padding: 40px;
      border-radius: 12px;
      box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }
    .modal-content { border-radius: 12px; }
    .spinner-border {
      width: 3rem;
      height: 3rem;
    }
  </style>
</head>
<body>
<div class="container">
  <h2 class="text-center">Asana Task Scheduler</h2>
  <form id="taskForm" method="post" enctype="multipart/form-data">
    <div class="mb-3"><label class="form-label">Excel File</label><input type="file" name="file" class="form-control" required></div>
    <div class="mb-3"><label class="form-label">Start Date</label><input type="date" name="start_date" class="form-control" required></div>
    <div class="mb-3">
      <label class="form-label">Workdays</label><br>
      {% for day in ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'] %}
        <div class="form-check form-check-inline">
          <input class="form-check-input" type="checkbox" name="workdays" value="{{day}}" {% if day in ['Monday','Tuesday','Wednesday','Thursday','Friday'] %}checked{% endif %}>
          <label class="form-check-label">{{day}}</label>
        </div>
      {% endfor %}
    </div>
    <div class="mb-3"><label class="form-label">Work Hours per Day</label><input type="number" name="work_hours" class="form-control" value="8.5" step="0.1" required></div>
    <div class="mb-3"><label class="form-label">Parent Task ID</label><input type="text" name="parent_task_id" class="form-control" required></div>
    <div class="mb-3"><label class="form-label">Assignee Email</label><input type="email" name="assignee" class="form-control" required></div>
    <div class="text-center"><button type="submit" class="btn btn-primary btn-lg">Submit</button></div>
    <div id="progress-container" class="mt-4" style="display:none;" class="text-center">
      <label class="form-label">Processing Tasks...</label>
      <div class="d-flex justify-content-center mt-3">
        <div class="spinner-border text-primary" role="status">
          <span class="visually-hidden">Loading...</span>
        </div>
      </div>
    </div>
  </form>
</div>

<!-- Modal -->
<div class="modal fade" id="completionModal" tabindex="-1" aria-labelledby="completionModalLabel" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered">
    <div class="modal-content">
      <div class="modal-header"><h5 class="modal-title" id="completionModalLabel">Task Insertion Completed</h5></div>
      <div class="modal-body">All tasks have been successfully created in Asana.</div>
      <div class="modal-footer"><button type="button" class="btn btn-primary" onclick="window.location.href='/'">Back to Main Page</button></div>
    </div>
  </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script>
  document.getElementById("taskForm").addEventListener("submit", function(event) {
    event.preventDefault();
    const form = event.target;
    const formData = new FormData(form);
    document.getElementById("progress-container").style.display = 'block';

    const xhr = new XMLHttpRequest();
    xhr.open("POST", "/submit", true);
    xhr.responseType = "text";

    xhr.onreadystatechange = function () {
      if (xhr.readyState === 4) {
        if (xhr.responseText.includes("STATUS:success")) {
          const modal = new bootstrap.Modal(document.getElementById('completionModal'));
          modal.show();
        } else {
          alert("❌ Some or all tasks failed to be created in Asana. Please check your credentials and network.");
        }
      }
    };

    xhr.send(formData);
  });
</script>
</body>
</html>'''

@app.route('/', methods=['GET'])
def upload_file():
    return render_template_string(HTML_TEMPLATE)

@app.route('/submit', methods=['POST'])
def submit():
    try:
        print("[LOG] Starting task submission process")

        file = request.files['file']
        print(f"[LOG] Received file: {file.filename}")

        start_date = datetime.datetime.strptime(request.form['start_date'], '%Y-%m-%d')
        workdays = [{'Sunday': 6, 'Monday': 0, 'Tuesday': 1, 'Wednesday': 2, 'Thursday': 3, 'Friday': 4, 'Saturday': 5}[d] for d in request.form.getlist('workdays')]
        work_hours = float(request.form['work_hours'])
        parent_task_id = request.form['parent_task_id']
        assignee = request.form['assignee']
        print(f"[LOG] Params - Start Date: {start_date}, Work Hours: {work_hours}, Parent Task ID: {parent_task_id}, Assignee: {assignee}")

        df = pd.read_excel(file)
        print(f"[LOG] Excel data loaded. Rows: {len(df)}")
        df.columns = df.columns.str.strip().str.lower()
        df = df.rename(columns={'estimation of hours per task': 'hours', 'task number': 'task'})
        df['task'] = df['task'].astype(str).str.strip()
        df = df[df['task'].str.lower() != 'nan']

        print(f"[LOG] Tasks after filtering: {len(df)}")

        current_date = start_date
        hours_left = work_hours
        due_dates = []
        for idx, hours in enumerate(df['hours']):
            h = float(hours) if not pd.isna(hours) else 0
            if h == 0:
                due_dates.append(current_date.strftime('%Y-%m-%d'))
                continue
            while h > 0:
                if current_date.weekday() in workdays:
                    if h <= hours_left:
                        hours_left -= h
                        h = 0
                    else:
                        h -= hours_left
                        hours_left = 0
                if h > 0:
                    current_date += datetime.timedelta(days=1)
                    while current_date.weekday() not in workdays:
                        current_date += datetime.timedelta(days=1)
                    hours_left = work_hours
            due_dates.append(current_date.strftime('%Y-%m-%d'))

        df['due'] = due_dates
        print("[LOG] Due dates calculated")

        token = os.getenv('ASANA_PERSONAL_ACCESS_TOKEN')
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
        total = len(df)
        failures = 0
        failed_tasks_log = []
        log_filename = "failed_tasks_log.txt"

        @stream_with_context
        def generate():
            nonlocal failures
            for i in reversed(range(total)):
                task_name = df.iloc[i]['task']
                due_date = df.iloc[i]['due']
                payload = {"data": {"name": task_name, "assignee": assignee, "due_on": due_date}}
                print(f"[LOG] Creating task {i+1}/{total}: {task_name} due {due_date}")
                try:
                    response = requests.post(f"https://app.asana.com/api/1.0/tasks/{parent_task_id}/subtasks", headers=headers, json=payload)
                    print(f"[LOG] Response Code: {response.status_code}, Body: {response.text}")
                    if response.status_code != 201:
                        failures += 1
                        failed_tasks_log.append(f"{task_name} | {due_date} | Status: {response.status_code} | Body: {response.text}\n")
                except Exception as api_error:
                    print(f"[ERROR] Failed to create task {task_name}: {str(api_error)}")
                    failed_tasks_log.append(f"{task_name} | {due_date} | Exception: {str(api_error)}\n")
                    failures += 1

                percent = int(((i+1)/total) * 100)
                yield f"data: PROGRESS:{percent}%\n\n"
                time.sleep(0.05)

            if failures > 0:
                with open(log_filename, "w", encoding="utf-8") as f:
                    f.writelines(failed_tasks_log)

            yield f"data: STATUS:{'success' if failures == 0 else 'fail'}\n\n"
            print("[LOG] Submission process completed")

        return Response(generate(), mimetype='text/event-stream')

    except Exception as e:
        print("[FATAL ERROR]", traceback.format_exc())
        return Response(f"data: STATUS:fail\n\nERROR:{str(e)}", mimetype='text/event-stream')

@app.route('/download-log')
def download_log():
    return send_file("failed_tasks_log.txt", as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)

import pandas as pd
import datetime
from flask import Flask, request, render_template_string
import os
from dotenv import load_dotenv
import requests

app = Flask(__name__)
load_dotenv()

HTML_TEMPLATE = '''
<!doctype html>
<title>Asana Task Scheduler</title>
<h2>Upload Task Excel File</h2>
<form method=post enctype=multipart/form-data>
  Excel File: <input type=file name=file><br><br>
  Start Date: <input type=date name=start_date><br><br>
  Workdays:<br>
  {% for day in ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'] %}
    <input type="checkbox" name="workdays" value="{{day}}" {% if day in ['Monday','Tuesday','Wednesday','Thursday','Friday'] %}checked{% endif %}>{{day}}<br>
  {% endfor %}
  <br>
  Work Hours per Day: <input type=number name=work_hours value=8.5 step=0.1><br><br>
  Parent Task ID: <input type=text name=parent_task_id><br><br>
  Assignee Email: <input type=email name=assignee><br><br>
  <input type=submit value=Submit>
</form>
'''

DAYS_MAP = {
    'Sunday': 6,
    'Monday': 0,
    'Tuesday': 1,
    'Wednesday': 2,
    'Thursday': 3,
    'Friday': 4,
    'Saturday': 5
}

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        try:
            if 'file' not in request.files:
                return "<h3>Error: No file uploaded.</h3>", 400

            file = request.files['file']
            if file.filename == '':
                return "<h3>Error: No file selected.</h3>", 400

            start_date_str = request.form['start_date'].strip()
            if not start_date_str:
                return "<h3>Error: Start date is required.</h3>", 400

            start_date = datetime.datetime.strptime(start_date_str, '%Y-%m-%d')
            workdays = request.form.getlist('workdays')
            if not workdays:
                return "<h3>Error: At least one workday must be selected.</h3>", 400

            workdays = [DAYS_MAP[d] for d in workdays]
            work_hours = float(request.form['work_hours'])
            parent_task_id = request.form['parent_task_id'].strip()
            assignee_email = request.form['assignee'].strip()

            if not parent_task_id:
                return "<h3>Error: Parent Task ID is required.</h3>", 400
            if not assignee_email:
                return "<h3>Error: Assignee email is required.</h3>", 400

            try:
                df = pd.read_excel(file)
            except Exception as e:
                return f"<h3>Error reading Excel file: {str(e)}</h3>", 400

            df.columns = df.columns.str.strip().str.lower()
            df = df.rename(columns={
                'h of hours per task': 'estimation of hours per task',
                'task number': 'task number'
            })

            required_columns = ['task number', 'estimation of hours per task']
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                return f"<h3>Error: Missing required columns: {', '.join(missing_columns)}</h3>", 400

            df['task number'] = df['task number'].astype(str).str.strip()
            df = df[df['task number'].str.lower() != 'nan']

            def calculate_due_dates(start_date, task_hours):
                current_date = start_date
                due_dates = []
                hours_remaining_today = work_hours

                for hours in task_hours:
                    if pd.isna(hours) or hours <= 0:
                        due_dates.append(current_date.strftime("%Y-%m-%d"))
                        continue

                    hours = float(hours)
                    while hours > 0:
                        if current_date.weekday() in [4, 5]:
                            current_date += datetime.timedelta(days=1)
                            hours_remaining_today = work_hours
                            continue

                        if hours <= hours_remaining_today:
                            hours_remaining_today -= hours
                            hours = 0
                        else:
                            hours -= hours_remaining_today
                            current_date += datetime.timedelta(days=1)
                            while current_date.weekday() in [4, 5]:
                                current_date += datetime.timedelta(days=1)
                            hours_remaining_today = work_hours

                    due_dates.append(current_date.strftime("%Y-%m-%d"))

                return due_dates

            df["due date"] = calculate_due_dates(start_date, df["estimation of hours per task"].tolist())

            def version_key(val):
                return [str(part).zfill(5) for part in str(val).replace(' ', '').split('.') if part]

            df = df.sort_values(by="task number", key=lambda col: col.map(version_key)).reset_index(drop=True)

            asana_token = os.getenv("ASANA_PERSONAL_ACCESS_TOKEN")
            headers = {
                "Authorization": f"Bearer {asana_token}",
                "Content-Type": "application/json"
            }

            created_tasks = 0
            errors = []

            for _, row in df[::-1].iterrows():
                task_name = row.get("task number", "").strip()
                due_date = row.get("due date")

                if not task_name or not due_date:
                    errors.append(f"❌ Skipped row with missing task name or due date: {row.to_dict()}")
                    continue

                payload = {
                    "data": {
                        "name": task_name,
                        "assignee": assignee_email,
                        "due_on": due_date
                    }
                }

                response = requests.post(
                    f"https://app.asana.com/api/1.0/tasks/{parent_task_id}/subtasks",
                    headers=headers,
                    json=payload
                )

                if response.status_code == 201:
                    created_tasks += 1
                else:
                    errors.append(f"❌ Failed to create subtask {task_name}: {response.text}")

            result_html = f"<h2>Task Creation Results</h2>"
            result_html += f"<p><strong>Successfully created: {created_tasks} tasks</strong></p>"

            if errors:
                result_html += f"<p><strong>Errors: {len(errors)}</strong></p><ul>"
                for error in errors[:5]:
                    result_html += f"<li style='color:red'>{error}</li>"
                if len(errors) > 5:
                    result_html += f"<li>...and {len(errors)-5} more errors</li>"
                result_html += "</ul>"

            result_html += "<h3>Task Schedule:</h3>"
            result_html += df.to_html(index=False)
            result_html += "<br><br><a href='/'>Create More Tasks</a>"
            return result_html

        except Exception as e:
            return f"<h3 style='color:red;'>An error occurred: {str(e)}</h3>", 500

    return render_template_string(HTML_TEMPLATE)

if __name__ == '__main__':
    app.run(debug=True)

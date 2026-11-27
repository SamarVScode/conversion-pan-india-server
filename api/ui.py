from fastapi import APIRouter
from fastapi.responses import HTMLResponse

router = APIRouter()

HTML_TEMPLATE = """<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>Conversion Report Stream Server — Test Bench</title>
  <style>
    body { font-family: system-ui, sans-serif; background: #0f172a; color: #f8fafc; padding: 40px; }
    .card { background: #1e293b; padding: 24px; border-radius: 12px; max-width: 600px; margin: 0 auto; box-shadow: 0 10px 25px rgba(0,0,0,0.5); }
    h2 { margin-top: 0; color: #38bdf8; }
    .badge { background: #0369a1; padding: 4px 8px; border-radius: 4px; font-size: 12px; vertical-align: middle; }
    input[type="text"], select { width: 100%; padding: 12px; margin: 8px 0 16px; border-radius: 6px; border: 1px solid #475569; background: #0f172a; color: #fff; box-sizing: border-box; }
    button { background: #0284c7; color: white; border: none; padding: 12px 20px; border-radius: 6px; cursor: pointer; font-weight: bold; width: 100%; font-size: 15px; }
    button:hover { background: #0369a1; }
    .status { margin-top: 20px; padding: 12px; border-radius: 6px; background: #334155; font-family: monospace; }
  </style>
</head>
<body>
  <div class="card">
    <h2>Conversion Report Server <span class="badge">Standalone</span></h2>
    <label>Google Drive URL / ID:</label>
    <input type="text" id="driveUrl" placeholder="https://drive.google.com/file/d/...">
    <label>Report Sub-Type:</label>
    <select id="subType">
      <option value="sameday">Sameday (Sameday DC_View &amp; Sameday Agent_View)</option>
      <option value="d1">D-1 (D-1 DC_View &amp; D-1 Agent_View)</option>
    </select>
    <label>API Key:</label>
    <input type="text" id="apiKey" value="OoV81VZ6ugIQ5qu_JNKfDM0jEp0SQyhpuZMaPTv5BbQ">
    <button onclick="triggerJob()">Run Conversion Report Job</button>
    <div class="status" id="statusBox">Idle</div>
  </div>

  <script>
    async function triggerJob() {
      const url = document.getElementById('driveUrl').value;
      const subType = document.getElementById('subType').value;
      const apiKey = document.getElementById('apiKey').value;
      const box = document.getElementById('statusBox');

      if (!url) { alert('Please enter a Google Drive URL or File ID'); return; }
      box.textContent = 'Triggering conversion job...';

      try {
        const res = await fetch(`/convert-async?drive_url=${encodeURIComponent(url)}&sub_type=${encodeURIComponent(subType)}`, {
          headers: { 'X-API-KEY': apiKey }
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.detail || 'Request failed');

        box.textContent = `Job Started! ID: ${data.job_id}. Polling status...`;
        pollJob(data.job_id, apiKey);
      } catch (err) {
        box.textContent = `Error: ${err.message}`;
      }
    }

    async function pollJob(jobId, apiKey) {
      const box = document.getElementById('statusBox');
      const interval = setInterval(async () => {
        try {
          const res = await fetch(`/job/${jobId}`, { headers: { 'X-API-KEY': apiKey } });
          const data = await res.json();
          box.textContent = `Status: ${data.status} | Progress: ${data.progress || ''}`;

          if (data.status === 'done') {
            clearInterval(interval);
            box.innerHTML = `Complete! <a href="/job/${jobId}/result" style="color:#38bdf8; font-weight:bold;" target="_blank">Download Result (.xlsx)</a>`;
          } else if (data.status === 'error') {
            clearInterval(interval);
            box.textContent = `Job Failed: ${data.error}`;
          }
        } catch (e) {
          box.textContent = `Polling Error: ${e.message}`;
        }
      }, 3000);
    }
  </script>
</body>
</html>
"""


@router.get("/test", response_class=HTMLResponse)
async def test_ui():
    return HTMLResponse(content=HTML_TEMPLATE, status_code=200)

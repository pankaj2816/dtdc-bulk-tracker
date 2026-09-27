# 🚀 DTDC Bulk Consignment Tracker

An automated web dashboard to track DTDC consignment / docket serial numbers with AI-driven background captcha bypass, live progress streaming, and Excel export.

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy)

---

### 🌟 Key Features

1. **Dual Input Methods**:
   - **Upload Excel / CSV**: Drag & drop your daily dispatch sheet (`.xlsx`, `.xls`, `.csv`). The app features a live upload progress bar (MB & percentage) and automatically scans all sheets (`MIA-2`, `DAILY ENTRY`, etc.) to group dockets by date.
   - **Direct Paste**: Paste multiple serial numbers separated by commas, spaces, or line breaks.

2. **Automated AI-OCR Captcha Bypass**:
   - Integrated lightweight OCR (`ddddocr`) solves DTDC security captchas automatically in background batches (~0.5s).
   - Zero manual captcha typing required.

3. **Safe Server Chunking & Real-Time Streaming**:
   - Automatically chunks requests into safe batches of **25 parcels** to prevent DTDC server crashes (`HTTP 500`).
   - Streams live progress (`SSE - Server-Sent Events`) parcel-by-parcel into the interactive web UI.

4. **Rich Status Metrics & Filters**:
   - Live metrics: Total, Delivered, In Transit, Out for Delivery, Issues.
   - Filter by status tab or search by city, docket number, or destination in real-time.

5. **Dual Excel Export Options**:
   - ⚡ **Export Summary (.xlsx)**: Instant (< 0.1s) download of all tracked parcels with live status, EDD, route, checkpoints, and green highlights for delivered shipments.
   - 📁 **Merge Original (.xlsx)**: Updates the original multi-sheet workbook, adding `Live Status`, `EDD`, `Latest Activity`, and `Activity Date/Time` columns and highlighting delivered docket numbers in green (`#C6EFCE`).

---

### 💻 Running Locally

#### Option 1: Double-Click
Simply double-click **[`run.bat`](file:///d:/Work/lokesh_dtdc/run.bat)**. It will launch the application and open your browser at `http://127.0.0.1:5000`.

#### Option 2: Command Line
```bash
# Clone the repository
git clone https://github.com/pankaj2816/dtdc-bulk-tracker.git
cd dtdc-bulk-tracker

# Install dependencies
pip install -r requirements.txt

# Run the app
python app.py
```
Open your browser at **`http://localhost:5000`**.

---

### ☁️ Free Cloud Deployment (Render / Railway / Hugging Face)

#### Deploy to Render (Recommended - Free):
1. Create a free account at [render.com](https://render.com).
2. Click **New +** > **Web Service**.
3. Connect your GitHub repository `pankaj2816/dtdc-bulk-tracker`.
4. Render will automatically detect `render.yaml` or set:
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --threads 4 --timeout 120`
5. Click **Create Web Service** to get a public live URL (e.g. `https://dtdc-bulk-tracker.onrender.com`).

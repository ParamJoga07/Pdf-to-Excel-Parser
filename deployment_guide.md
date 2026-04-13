# Deployment Guide: Hosting Your PDF to Excel App

Your application is ready for deployment! Because it uses Python (FastAPI) and requires system libraries for PDF processing, it is best hosted on platforms that support **Docker** or **Python Web Services**.

---

## 1. Prepare Your Project
Before hosting, ensure all files are in your project folder:
- `main.py` (Backend)
- `pdf_to_excel.py` (Processing logic)
- `static/` (Frontend UI)
- `requirements.txt` (Dependencies)
- `Dockerfile` (Container configuration)

> [!TIP]
> Initialize a Git repository and push your code to **GitHub**. Most hosting providers connect directly to GitHub for automatic deployments.

---

## 2. Recommended Hosting Platforms

### Option A: Render (Best for Free/Easy)
Render is very popular for its simplicity and reliable free tier.
1. Create a free account at [render.com](https://render.com/).
2. Click **New +** > **Web Service**.
3. Connect your GitHub repository.
4. Render should automatically detect your environment. Use these settings if prompted:
   - **Runtime**: `Docker` (Strongly recommended as it handles all PDF library dependencies correctly).
   - **Build Command**: (Leave blank if using Docker)
   - **Start Command**: (Leave blank if using Docker)
5. Click **Deploy Web Service**.

### Option B: Railway (Fastest Setup)
Railway is excellent for speed and "it just works."
1. Sign up at [railway.app](https://railway.app/).
2. Click **New Project** > **Deploy from GitHub repo**.
3. Select your repository.
4. Railway will see the `Dockerfile` and deploy everything automatically.

### Option C: PythonAnywhere (Python-Centric)
If you prefer a purely Python environment:
1. Use [PythonAnywhere](https://www.pythonanywhere.com/).
2. You will need to set up a "Web App" manually using their FastAPI/WSGI instructions. *Note: Docker is easier for this specific project.*

---

## 3. Important Notes for Deployment

> [!IMPORTANT]
> **Processing Time**: Converting 100 PDFs can take a minute or two. Ensure your hosting provider's "Timeout" setting is high enough (usually 30-60 seconds is the default; Render/Railway are generous with this).

> [!NOTE]
> **Ephemeral Storage**: Your app uses temporary folders for processing. This is safe, as the folders are cleaned up after each conversion to prevent the server from running out of space.

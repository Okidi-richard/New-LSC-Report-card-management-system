# How to Put the Report Card System Online (Phone-Friendly)

Once deployed, **any phone** can open the system in Chrome and generate report cards.
No computer is needed for daily use.

This guide uses **Render** (free tier). It takes about 10–15 minutes the first time.

---

## What you need

- A free account on [https://render.com](https://render.com)
- A free account on [https://github.com](https://github.com) (to store the code)
- The `uganda_report_card_system` folder on a computer (one-time only)

---

## Step 1 – Put the code on GitHub

1. Go to [https://github.com](https://github.com) and sign up / log in.
2. Click the **+** (top right) → **New repository**.
3. Name it: `uganda-report-card`
4. Keep it **Public**.
5. Click **Create repository**.
6. On your computer, open the `uganda_report_card_system` folder.
7. Upload all the files to the new GitHub repository  
   (easiest way: use the GitHub website “uploading an existing file” or GitHub Desktop).

You should see these important files in the repository:
- `requirements.txt`
- `Procfile`
- `runtime.txt`
- `run_app.py`
- `report_card_system.py`
- `app/` folder (with `app.py` and `templates/`)

---

## Step 2 – Deploy on Render

1. Go to [https://render.com](https://render.com) and sign up (you can use your GitHub account).
2. Click **New +** → **Web Service**.
3. Connect your GitHub account if asked, then select the repository `uganda-report-card`.
4. Fill in the settings:

| Setting              | Value                                      |
|----------------------|--------------------------------------------|
| Name                 | `uganda-report-card` (or any name)         |
| Region               | Choose the closest (e.g. Frankfurt)        |
| Branch               | `main`                                     |
| Runtime              | Python                                     |
| Build Command        | `pip install -r requirements.txt`          |
| Start Command        | `gunicorn app.app:app --bind 0.0.0.0:$PORT --workers 1 --timeout 120` |
| Instance Type        | **Free**                                   |

5. Click **Create Web Service**.
6. Wait 3–8 minutes while Render builds and starts the app.
7. When the status becomes **Live**, you will see a URL like:

   `https://uganda-report-card.onrender.com`

---

## Step 3 – Use it from any phone

1. Open **Chrome** on your phone.
2. Type the Render URL (the one ending in `.onrender.com`).
3. You will see the same Report Card System interface.
4. Download the Excel template, fill it (or fill on a computer and upload from the phone), generate report cards, and download the ZIP.

**Tip:** Add the page to your phone Home Screen for quick access:
- Chrome → Menu (⋮) → “Add to Home screen”.

---

## Important notes about the Free plan

- The free service **sleeps** after about 15 minutes of no use.
- The first time you open it after sleeping, it may take 30–60 seconds to wake up. Just wait.
- For a school that generates report cards only a few times per term, the free plan is usually enough.
- If you need it always fast, you can upgrade later (paid plans start from a few dollars per month).

---

## Alternative free hosts

If Render does not work for you, the same files also work on:

- **Railway** → [https://railway.app](https://railway.app)
- **PythonAnywhere** → [https://www.pythonanywhere.com](https://www.pythonanywhere.com) (a bit more manual setup)

---

## After deployment – daily use (phone only)

1. Open the website on your phone.
2. Download the Excel template (or keep a filled copy in Google Drive / WhatsApp).
3. Fill student names and marks.
4. Upload the Excel → Generate → Download ZIP of all report cards.
5. Share the PDFs with parents or teachers via WhatsApp.

That’s it. The system is now accessible from any phone with internet.

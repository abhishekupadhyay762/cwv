# Core Web Vitals Bulk Checker

A Streamlit app that checks real-user Core Web Vitals (Chrome UX Report field data) and Lighthouse lab scores for up to 200 URLs at once, on mobile and desktop, using Google's PageSpeed Insights API.

There is no password. Each visitor pastes their own free PageSpeed Insights API key, so every check runs on the visitor's own Google quota. Keys are kept only in the visitor's browser session and are never saved or logged.

## What it reports

For each URL and device:

- **Core Web Vitals assessment** (Passed / Failed / No data) from real Chrome users
- **Field data source**: URL-level, Origin (domain-wide fallback) or No data
- **Field metrics** (75th percentile, last 28 days): LCP, INP, CLS, FCP, TTFB
- **Lighthouse scores**: Performance, Accessibility, Best practices, SEO, Agentic browsing (passed of applicable audits)
- **Lab metrics**: LCP, FCP, TBT, CLS, Speed Index
- **Top fixes**: the three Lighthouse performance insights with the biggest estimated savings
- **Redirects**: the final URL when the tested URL redirects

Results show as colour-coded tables (Mobile, Desktop, Mobile vs desktop, Flags) and download as an Excel report or CSV. The optional Google Sheets sync reads URLs from column A of a sheet and writes results next to them.

## This is the team version

This app is separate from the personal, password-protected checker. It lives in its own folder, its own GitHub repository and its own Streamlit link. Don't copy these files into the personal app's folder or repository.

## Test locally (Windows)

1. Unzip `core-web-vitals-team.zip` into your `scratch` folder. It creates a new folder:
   ```
   C:\Users\Admin\.gemini\antigravity\scratch\core-web-vitals-team
   ```
2. Open Command Prompt in that folder:
   ```
   cd C:\Users\Admin\.gemini\antigravity\scratch\core-web-vitals-team
   ```
3. Install or update the packages (this doesn't affect the personal app):
   ```
   pip install --upgrade -r requirements.txt
   ```
4. Start the app:
   ```
   streamlit run app.py
   ```
5. Open **http://localhost:8501** in your browser, paste a PageSpeed Insights API key and run a few URLs.

Press `Ctrl + C` in Command Prompt to stop the app. If the personal app is already running, this one opens on **http://localhost:8502** instead.

## Deploy on Streamlit Community Cloud (as a new app)

1. On GitHub, create a **new repository**, for example `core-web-vitals-team`.
2. Upload `app.py`, `requirements.txt` and `README.md`.
3. Add the theme file: click **Add file > Create new file**, type `.streamlit/config.toml` as the file name (the slash creates the folder), paste in the contents of `config.toml` and commit.
4. Go to [share.streamlit.io](https://share.streamlit.io/), click **Create app**, choose the new repository, set the main file to `app.py`, pick a new app URL (for example `cwv-team-checker.streamlit.app`) and click **Deploy**.
5. No secrets are needed. Share the new link with your team; each person pastes their own API key.

The personal app (`core-web-vitals-score.streamlit.app`) and its repository and secrets stay exactly as they are.

## Getting a PageSpeed Insights API key

1. Go to the [PageSpeed Insights API get-started page](https://developers.google.com/speed/docs/insights/v5/get-started) and click **Get a Key**, or create a key under [Credentials in Google Cloud Console](https://console.cloud.google.com/apis/credentials).
2. Make sure the **PageSpeed Insights API** is enabled for the project.
3. Recommended restrictions: API restriction set to *PageSpeed Insights API* only; Application restriction set to *None* (a website/HTTP referrer restriction blocks server requests from this app).

## Google Sheets sync (optional)

Visitors use their own Google service account:

1. Enable the **Google Sheets API** in a Google Cloud project.
2. Create a service account, then **Keys > Add key > Create new key > JSON**.
3. Upload the .json file in the app, then share the Google Sheet with the service account email as **Editor**.
4. Put URLs in column A from row 2 and paste the sheet link into the app.

Only rows that get a result are written; rows without a URL are left untouched.

# Ember — Setup Guide

Three parts: a Google Sheet (database), an Apps Script Web App (backend/API), and a static page on GitHub Pages (frontend). Same pattern as your Clarity app.

## What's new in this version
- Renamed to **Ember** — a flame mark as the logo (used in the top bar, browser tab icon, and splash screen)
- A branded loading splash (breathing flame + dot loader) on app open, same idea as Clarity's
- The whole UI now **reacts to your goal**: colors, the flame gauge, and the copy shift automatically —
  warm amber/ember tones and "Building" language for **gain**, cool blue tones and "Trimming" language for **lose**, sage tones and "Holding" language for **maintain**. Change your goal in Profile and the whole app re-themes.
- The calorie ring is now a flame that fills from the bottom as you log food, instead of a plain progress ring
- Rounder, higher-contrast cards, gradient buttons, and meal-type icons in the log

## 1. Create the Google Sheet
1. Go to sheets.google.com → Blank spreadsheet. Name it "Ember DB".
2. Extensions → Apps Script. Delete any starter code in `Code.gs`.
3. Paste in the entire contents of `Code.gs` (provided alongside this guide).
4. Save the project (name it "Fuel Ledger Backend").
5. Run the function `ensureSheetsExist` once from the Apps Script editor (select it from the function dropdown, click Run). The first run will ask for permissions — approve them. This creates all the tabs (Profile, WeightLog, FoodDatabase, Pantry, FoodLog, MealPlans, Config) and seeds a starter food list.

## 2. Deploy the Web App
1. In the Apps Script editor: Deploy → New deployment.
2. Click the gear icon next to "Select type" → Web app.
3. Description: "Fuel Ledger API". Execute as: **Me**. Who has access: **Anyone**.
4. Click Deploy, authorize again if asked.
5. Copy the **Web app URL** — it looks like `https://script.google.com/macros/s/XXXXX/exec`.

Whenever you edit `Code.gs` later, use **Deploy → Manage deployments → edit (pencil) → New version** so the live URL picks up your changes — a plain save doesn't update the deployed version.

## 3. Wire up the frontend
1. Open `index.html`.
2. Find the line near the top of the `<script>` section:
   ```js
   const API_URL = "PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE";
   ```
3. Replace it with the URL you copied in step 2.

## 4. Host on GitHub Pages
1. Create a new GitHub repo (e.g. `fuel-ledger`).
2. Upload `index.html` to the repo root.
3. Repo Settings → Pages → Source: **Deploy from a branch** → branch `main`, folder `/root`. Save.
4. Your app will be live at `https://<your-username>.github.io/fuel-ledger/` within a minute or two.
5. Add it to your phone's home screen (Share → Add to Home Screen) for an app-like feel.

## 5. (Optional but recommended) Add your Gemini API key
1. Get a free key at aistudio.google.com/apikey (same place you likely got the one for Clarity — you can reuse a single key across both apps, or create a separate one to keep usage separate).
2. In the app: Profile tab → paste it into "Gemini API key" → Save.
3. Without a key, the app still works — meal plans fall back to a simple rule-based picker using your Food Database and pantry, and natural-language food logging (the "Describe" tab) won't be available until a key is set.

## 6. First run in the app
1. Open the app → Profile tab → fill in your stats and hit Save. This calculates your daily calorie/macro targets.
2. Add a few items to your Pantry (Plan tab).
3. Go to Plan → toggle "Office day" if applicable → Generate plan.
4. Log meals from Dashboard → Log (search your database, describe in plain language, or quick-add from pantry).

## Notes on the food database
The seed list has ~14 common items (roti, dal, rice, paneer, chicken, eggs, etc.) so search works immediately. Add more anytime — either directly in the `FoodDatabase` tab of the Sheet, or extend the app later with an "Add new food" form (the backend already has an `addFoodToDatabase` action ready for it).

## What's built vs. what's a good next step
**Built:** profile + BMR/TDEE/macro targets, food logging (search / natural-language AI / pantry quick-add), AI-personalized meal plans (with office-day and pantry-only modes, plus a non-AI fallback), pantry management, weight + calorie trend charts, saved plan history.

**Good next additions:** editable/undo on food log entries, a "copy yesterday's log" button, push-style daily reminder (browser notification), barcode or photo-based food logging, multi-day meal plan generation, exporting weight/calorie data to CSV.

# TDF Fantasy Mercato Manager

## Overview

This Python script automates the extraction of rider data from the official Tour de France Fantasy (by Tissot) API. It is specifically designed to help managers prepare for their private league's "Mercato" (auction draft).

The script generates a color-coded Excel spreadsheet containing all active riders, their costs, categories, and availability status. Crucially, it preserves your manual inputs ("Infos sup" and "Etoiles misées") if you run it multiple times to update the market data.

## Usage

Open the script (scrap_infos_riders.py) and ensure your AUTH_TOKEN is up to date (see the Annual Maintenance section below).

Run the script from your terminal:

```bash
python scrap_infos_riders.py
```

The script will generate (or update) a file named Mercato_TDF_Complet.xlsx in the same directory.

## Annual Maintenance (Updating for future seasons)

If you are reusing this script for a new Tour de France season, the API security and structure will likely have changed. You must verify and update the following variables in the CONFIGURATION section of the script:

1. The Authentication Token (AUTH_TOKEN) - Must be updated regularly

Tokens expire for security reasons. You will need to fetch a fresh one:

Log into the TDF Fantasy website on your browser.

Open Developer Tools (F12) -> Network tab.

Refresh the page and filter by XHR.

Click on a request (e.g., searchjoueurs?lg=fr) and check the Request Headers.

Copy the full string next to Authorization (it should start with Token eyJ...) and paste it into the AUTH_TOKEN variable in .env file.

2. The Access Key (X-Access-Key) - Changes every year

The Tissot developers use a custom header to prevent basic scraping.

In the same Request Headers where you found the token, look for X-Access-Key.

Update the 'X-Access-Key' value in the HEADERS dictionary in the script (e.g., '630@19.30@'). If you skip this, you will get a 401 Unauthorized error.

3. The API Endpoint (API_URL)

Check if the API version has changed. For example, if they upgrade from v1 to v2, the URL will change:

Old: https://fantasybytissot.letour.fr/v1/private/searchjoueurs?lg=fr

New: Update the API_URL variable accordingly.

4. Rider Categories (CATEGORY_COLORS)

The game occasionally changes the names of rider roles (e.g., from "Sprinter" to "Baroudeur"). If the script generates an Excel file with missing background colors:

Check your terminal output for warnings like: ⚠️ Catégorie inconnue...

Update the keys in the CATEGORY_COLORS dictionary to match the new exact terminology returned by the API.

## TODO

- ajouter les notes sur le coté pr pas les perdre
- programme qui calcule les meilleures stratégies en fonction de l'attribution des points


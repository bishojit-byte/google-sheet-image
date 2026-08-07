# Google Sheets Bulk Image Inserter (Apps Script)

A custom Google Apps Script project for Google Sheets that adds an interactive tool to bulk insert images into cells **horizontally** (left to right across columns) or **vertically** (top to bottom down rows).

---

## 🌟 Features
- **Custom Menu Button**: Adds a `⚡ Image Inserter` menu item directly in the Google Sheets menu bar.
- **Interactive Popup Modal Dialog**: Sleek, modern UI for pasting multiple image URLs.
- **Direction Options**:
  - ↔️ **Horizontal**: Inserts images across columns in the same row.
  - ↕️ **Vertical**: Inserts images down rows in the same column.
- **Image Types**:
  - **In-Cell Image (Native)**: Uses Google Apps Script `CellImageBuilder` to insert genuine in-cell images.
  - **`=IMAGE("url")` Formula**: Standard formula insertion fallback.
- **Auto Cell Sizing**: Automatically adjusts column width and row height to make inserted images visible.
- **Demo Mode**: Includes a "Demo URLs" button for easy testing with placeholder images.

---

## 📂 Project Structure

```
E:\Applications\Google Sheet Appscripts\
├── Code.gs             # Server-side Apps Script (onOpen menu & backend handler)
├── Dialog.html         # Custom popup HTML modal interface & styling
├── appsscript.json     # Manifest file with permissions & runtime settings
├── README.md           # Setup & GitHub repository guide
└── .gitignore          # Git ignore configuration
```

---

## 🚀 How to Install & Run in Google Sheets

### Method 1: Container-Bound Script (Quickest & Recommended)
1. Open any Google Sheet (or create a new one).
2. Click **Extensions** > **Apps Script** in the top menu.
3. In the Apps Script editor:
   - Replace the code in `Code.gs` with the content from [`Code.gs`](./Code.gs).
   - Click **+** next to **Files** > select **HTML**, name it `Dialog`, and paste the content from [`Dialog.html`](./Dialog.html).
   - (Optional) Enable manifest editing under **Project Settings** > check *Show "appsscript.json" manifest file in editor* and paste [`appsscript.json`](./appsscript.json).
4. Click the **Save** icon (💾) or press `Ctrl + S`.
5. Return to your Google Sheet and **refresh the webpage**.
6. You will see a new menu in the menu bar: `⚡ Image Inserter` > **Bulk Insert Images...**.

---

## 💻 How to Push to GitHub

To link this local codebase with a GitHub repository:

1. Open your terminal / PowerShell in `E:\Applications\Google Sheet Appscripts`.
2. Initialize git:
   ```bash
   git init
   git add .
   git commit -m "Initial commit: Bulk Image Inserter Google Apps Script"
   ```
3. Create a new repository on GitHub (e.g. `google-sheet-image-inserter`).
4. Link the remote repository and push:
   ```bash
   git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
   git branch -M main
   git push -u origin main
   ```

---

## 🛠️ Usage Instructions

1. Click on any cell in your sheet where you want image insertion to start (e.g. `B2`).
2. Click **`⚡ Image Inserter`** > **`Bulk Insert Images...`**.
3. Paste image URLs (one per line).
4. Select direction:
   - ↔️ **Horizontal (Row)** to fill cells to the right.
   - ↕️ **Vertical (Column)** to fill cells downwards.
5. Click **Insert Images**.

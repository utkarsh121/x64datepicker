# 🗓️ Custom Date Picker for Microsoft Excel

This is a simple, modern Task Pane Add-in for Microsoft Excel that provides a calendar interface for quickly selecting and inserting dates into the currently active cell. This project is hosted entirely on **GitHub Pages** for easy, free distribution and uses the secure Office JavaScript API for cross-platform compatibility.

## ✨ Features

* **Quick Insertion:** Select any date in the calendar to insert the date value directly into the selected Excel cell.
* **Automatic Formatting:** The selected cell is automatically formatted to the standard date style (`m/d/yyyy`).
* **Intuitive Navigation:** Navigate month-by-month and year-by-year using the built-in controls.
* **Real-Time Status:** Provides instant feedback showing the inserted date and the target cell address.

## 🚀 Installation & Sharing (Sideloading)

This add-in is shared using a mechanism called "sideloading," which requires a single file: the `manifest.xml`. Your friends only need this file to install the add-in—they do not need to install Node, NPM, or access any local web server.

### Step 1: Get the Manifest File

1.  Go to the main repository page: **[https://github.com/utkarsh121/x64datepicker](https://github.com/utkarsh121/x64datepicker)**
2.  Locate and download the **`manifest.xml`** file.

### Step 2: Sideload in Excel

1.  Open **Microsoft Excel** (Desktop Application).
2.  Go to the **Insert** tab on the ribbon.
3.  Click the **Get Add-ins** button (or "My Add-ins").
4.  In the Add-ins window, look for and click the link labeled **"Manage My Add-ins"**.
5.  Select **"Upload My File..."**
6.  Select the **`manifest.xml`** file you downloaded in Step 1.
7.  Click **Upload**.

### Step 3: Launch the Date Picker

1.  The add-in will now appear in your Excel ribbon, usually on the **Home** tab, in a group called **Date Picker Tools**.
2.  Click the **"Date Picker"** button to open the calendar panel on the right and begin inserting dates.

## ⚙️ Development & Hosting Details

The web content for this add-in is hosted live at: **[https://utkarsh121.github.io/x64datepicker/](https://utkarsh121.github.io/x64datepicker/)**. The `manifest.xml` points to this public, HTTPS-enabled URL.

### Core Technologies

* **Office JavaScript API:** The standard library for interacting with Excel objects.
* **Pure JavaScript:** Handles all calendar calculations and front-end logic.
* **Office UI Fabric Core:** Used for consistent styling and Microsoft Fluent UI icons.

## 📄 License

This project is licensed under the **MIT License**.

# SR&ED Projects Outlook Add-in

This Outlook add-in provides quick access to the SR&ED Project Intake Form for internal projects.

## ⚠️ Important Notice
This add-in is intended **only for internal SR&ED projects**. For client projects, please follow the standard Business Development (BD) process instead.

## Files Required

1. **index.html** - The main interface file
2. **manifest.xml** - The add-in configuration file
3. **icon-32.png** - Small icon (32x32 pixels)
4. **icon-80.png** - Large icon (80x80 pixels)

## Setup Instructions

### Step 1: Update the Form URL
Before deploying, you must update the Microsoft Forms URL in `index.html`:

1. Open `index.html`
2. Find this line (around line 114):
   ```javascript
   const FORM_URL = "https://forms.office.com/Pages/ResponsePage.aspx?id=YOUR_FORM_ID_HERE";
   ```
3. Replace `YOUR_FORM_ID_HERE` with your actual SR&ED Project Intake Form URL

### Step 2: Upload Files to GitHub
Upload these files to your GitHub repository:
- Path: `ep18432-form-addon/SHREDProjects/`
- Files: `index.html`, `manifest.xml`, `icon-32.png`, `icon-80.png`

### Step 3: Deploy to Outlook

#### Option A: Individual Installation (Sideload)
1. Open Outlook (desktop or web)
2. Go to **Get Add-ins** or **Manage Add-ins**
3. Click **My add-ins** → **Add a custom add-in** → **Add from URL**
4. Enter the manifest URL: `https://eastpoint-engineering.github.io/ep18432-form-addon/SHREDProjects/manifest.xml`

#### Option B: Organization-Wide Deployment (Microsoft 365 Admin)
1. Go to Microsoft 365 Admin Center
2. Navigate to **Settings** → **Integrated apps**
3. Click **Upload custom apps**
4. Upload the `manifest.xml` file
5. Configure deployment settings for your organization

## Customization

### Change the Icon
Replace `icon-32.png` and `icon-80.png` with your own icons. Requirements:
- icon-32.png: 32x32 pixels
- icon-80.png: 80x80 pixels
- Format: PNG with transparent background recommended

### Modify the Interface
Edit `index.html` to change:
- Colors (search for `#0078d4` to update the blue theme)
- Text and messaging
- Button labels
- Layout and styling

### Update the Add-in ID
If you need a unique ID (to run alongside other add-ins):
1. Generate a new GUID at https://guidgenerator.com/
2. Replace the `<Id>` value in `manifest.xml`

## Troubleshooting

### Add-in doesn't appear in Outlook
- Verify the manifest URL is accessible
- Check that GitHub Pages is enabled for your repository
- Wait a few minutes for Outlook to cache the add-in

### Form doesn't open
- Verify the form URL is correct in `index.html`
- Check that the form permissions allow access from your organization
- Test the form URL directly in a browser

### Changes not reflected
- Clear Outlook cache
- Remove and re-add the add-in
- Hard refresh the browser (Ctrl+F5) if using Outlook web

## Support
For issues or questions, contact your IT department or visit the GitHub repository support page.

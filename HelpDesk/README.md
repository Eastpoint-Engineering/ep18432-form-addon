# Help Desk Outlook Add-in
**Cross-Platform Support: Desktop, Web, and Mobile**

This Outlook add-in provides quick access to submit help desk tickets via email across all Outlook platforms.

## Platform Support

✅ **Outlook Desktop** (Windows & Mac)  
✅ **Outlook Web** (OWA / outlook.office.com)  
✅ **Outlook Mobile** (iOS & Android)  
✅ **Outlook on iPad**

## Files Required

1. **index.html** - The main interface file (cross-platform compatible)
2. **manifest.xml** - The add-in configuration file (multi-platform settings)
3. **icon-32.png** - Small icon (32x32 pixels)
4. **icon-80.png** - Large icon (80x80 pixels)

## Key Changes for Cross-Platform Support

### Manifest Updates
- Added `<TabletSettings>` for web and iPad support
- Added `<PhoneSettings>` for mobile support
- Supports both `ItemRead` and `ItemEdit` modes
- Works in compose and read modes

### HTML/JavaScript Updates
- Mobile-responsive design with media queries
- Multiple fallback methods for email creation
- Platform detection and appropriate API usage
- Mailto link fallback for maximum compatibility

## Setup Instructions

### Step 1: Upload Files to GitHub

Upload these files to your GitHub repository:
- **Path:** `ep18432-form-addon/HelpDesk/`
- **Files:** `index.html`, `manifest.xml`, `icon-32.png`, `icon-80.png`

Make sure GitHub Pages is enabled for your repository.

### Step 2: Deploy to Outlook

#### Option A: Organization-Wide Deployment (Recommended)
**For deployment across all platforms including mobile:**

1. Go to **Microsoft 365 Admin Center** (admin.microsoft.com)
2. Navigate to **Settings** → **Integrated apps**
3. Click **Upload custom apps**
4. Upload the `manifest.xml` file
5. Configure deployment settings:
   - Select user groups or entire organization
   - Enable for all platforms (Desktop, Web, Mobile)
6. Click **Deploy**

**Note:** Organization-wide deployment is required for mobile support. Individual sideloading is not available on mobile devices.

#### Option B: Individual Installation (Desktop & Web Only)
**This method does NOT work for mobile devices:**

1. Open Outlook (desktop or web)
2. Go to **Get Add-ins** or **Manage Add-ins**
3. Click **My add-ins** → **Add a custom add-in** → **Add from URL**
4. Enter: `https://eastpoint-engineering.github.io/ep18432-form-addon/HelpDesk/manifest.xml`

### Step 3: Verify Deployment

#### Desktop (Windows/Mac)
1. Open Outlook desktop app
2. Click on any email
3. Look for "Help Desk" in the ribbon or add-in panel
4. Click to test

#### Web (outlook.office.com)
1. Go to outlook.office.com
2. Click on any email
3. Look for "Help Desk" button in the ribbon or "..." menu
4. Click to test

#### Mobile (iOS/Android)
1. Open Outlook mobile app
2. Tap on any email
3. Swipe up to see add-ins or tap "..." menu
4. Look for "Help Desk"
5. Tap to test

**Important:** Mobile access requires organization-wide deployment through Microsoft 365 Admin Center.

## Configuration

### Email Recipients
The add-in is preconfigured with:
- **To:** help@impsolutions.com
- **CC:** Kaelan.workman@eastpoint.ca

To change these, edit `index.html` lines 125-126:
```javascript
const TO_RECIPIENT = "help@impsolutions.com";
const CC_RECIPIENT = "Kaelan.workman@eastpoint.ca";
```

### Visual Customization

#### Change Colors
Edit the CSS in `index.html`:
- Primary color: Search for `#0078d4` (Microsoft blue)
- Replace with your brand color

#### Update Icons
Replace the icon files with your own:
- **icon-32.png:** 32x32 pixels (for ribbon/toolbar)
- **icon-80.png:** 80x80 pixels (for high-DPI displays)
- Format: PNG with transparent background recommended
- Keep the same filenames

#### Modify Text
Edit `index.html`:
- Title: Line 6 and line 95
- Description: Line 96
- Button text: Line 99

## Troubleshooting

### Add-in doesn't appear

**Desktop/Web:**
- Verify manifest URL is accessible
- Check GitHub Pages is enabled
- Clear Outlook cache and restart
- Wait 5-10 minutes for deployment

**Mobile:**
- Verify organization-wide deployment is complete
- Check that mobile platforms are enabled in deployment settings
- Restart Outlook mobile app
- Sign out and sign back in

### Button doesn't work

**All platforms:**
- Check browser/app console for errors (F12 on web)
- Verify Office.js library loads properly
- Test the mailto fallback manually

**Mobile specific:**
- Ensure proper network connectivity
- Update Outlook app to latest version
- Check that ReadWriteMailbox permission is granted

### Add-in works on desktop but not mobile
- Mobile requires **organization-wide deployment**
- Individual sideloading doesn't work on mobile
- Contact your Microsoft 365 admin to deploy organization-wide

### Email doesn't open on mobile
- The add-in uses multiple fallback methods
- If all else fails, it opens the device's default email client
- Ensure Outlook is set as default email app for best results

## Platform-Specific Notes

### Desktop
- Full API support
- Best user experience
- Can sideload for testing

### Web (Outlook on the web)
- Nearly identical to desktop functionality
- Same API availability
- Can sideload for testing

### Mobile (iOS/Android)
- Requires organization-wide deployment
- Limited screen space - optimized UI
- May open in device email client as fallback
- Cannot be sideloaded individually

### iPad
- Treated as tablet by Outlook
- Uses tablet settings from manifest
- Full feature support
- Organization-wide deployment recommended

## Security & Permissions

The add-in requires **ReadWriteMailbox** permission to:
- Create new email messages
- Set recipients (To, CC)
- Access mailbox context

This is a standard permission for email-creation add-ins and is approved by Microsoft.

## Best Practices

1. **Test on all platforms** before organization-wide deployment
2. **Use organization deployment** for consistent experience across devices
3. **Monitor user feedback** through support channels
4. **Keep icons simple** and recognizable at small sizes
5. **Test offline behavior** on mobile devices

## Support

For issues or questions:
- **IT Support:** Contact your Microsoft 365 administrator
- **Technical Issues:** Check GitHub repository issues
- **User Guide:** Share this README with end users

## Version History

- **v1.0.0.2** - Cross-platform support (Desktop, Web, Mobile)
- **v1.0.0.1** - Initial desktop-only release

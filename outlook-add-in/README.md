# PAPI Pass Creator - Outlook Add-in

Create PAPI passes directly from Outlook emails with a single click. This add-in integrates seamlessly with your PAPI system to streamline pass creation workflows.

## 🎯 Features

- **Quick Pass Creation**: Create a pass with just a phone number
- **Advanced Mode**: Create passes with all available fields (ID, areas, security levels, etc.)
- **Secure Configuration**: One-time setup with encrypted credential storage
- **Email Integration**: Auto-fills sender information from emails
- **No Backend Changes**: Works with your existing PAPI API

## 📋 Prerequisites

- Outlook Desktop (Windows/Mac) or Outlook on the web
- PAPI instance with API access
- Client ID and API Token from PAPI
- Node.js 14+ (for development/testing)

## 🚀 Installation

### Step 1: Clone or Extract Files

Ensure all files are in the `outlook-add-in` directory:

```
outlook-add-in/
├── manifest.xml
├── package.json
├── README.md
└── src/
    ├── commands/
    │   ├── commands.html
    │   └── commands.js
    ├── setup/
    │   ├── setup.html
    │   └── setup.js
    ├── shared/
    │   └── utils.js
    └── taskpane/
        ├── taskpane.html
        ├── taskpane.js
        ├── advanced.html
        └── advanced.js
```

### Step 2: Install Dependencies

```powershell
cd outlook-add-in
npm install
```

### Step 3: Update Manifest (Important!)

Edit `manifest.xml` and update:

1. **Add-in ID** (line 10): Generate a new GUID

   ```xml
   <Id>YOUR-UNIQUE-GUID-HERE</Id>
   ```

   Generate at: https://guidgenerator.com/

2. **URLs**: Replace `https://localhost:3000` with your hosting URL (or keep localhost for testing)

3. **Support URL** (line 18): Add your support page URL

### Step 4: Create Icon Assets

Create a folder `assets` in the `outlook-add-in` directory and add icon files:

- `icon-16.png` (16x16)
- `icon-32.png` (32x32)
- `icon-80.png` (80x80)

Or use placeholder icons for testing.

### Step 5: Start the Development Server

```powershell
npm run serve
```

This starts a local server at `http://localhost:3000`

### Step 6: Sideload the Add-in

#### For Outlook Desktop (Windows):

1. Open Outlook
2. Go to **File** → **Get Add-ins**
3. Select **My add-ins** tab
4. Under **Custom add-ins**, click **Add a custom add-in** → **Add from file**
5. Browse to `manifest.xml` and select it
6. Click **Install**

#### For Outlook on the Web:

1. Go to Outlook.com or office.com/mail
2. Click **Settings** (gear icon) → **View all Outlook settings**
3. Go to **Mail** → **Customize actions** → **Get Add-ins**
4. Select **My add-ins** → **Add a custom add-in** → **Add from URL**
5. Paste the URL to your `manifest.xml` file

#### For Testing with Office Developer Tools:

```powershell
npm run start
```

This automatically sideloads the add-in in Outlook Desktop.

## ⚙️ Configuration

### First-Time Setup

1. Open any email in Outlook
2. Click the **⚙️ Setup** button in the ribbon
3. Fill in your PAPI credentials:
   - **API Base URL**: `https://your-papi-domain.com` (without `/passes`)
   - **Client ID**: Your organization's client ID (e.g., `123`)
   - **API Token**: Your secure API token from PAPI
   - **Default Event ID**: The event ID for pass creation (e.g., `456`)
4. Click **Test Connection** to verify
5. Click **Save Configuration**

### Configuration Details

All credentials are stored securely using Office.js Roaming Settings, which:

- Are encrypted by Microsoft
- Sync across devices
- Are user-specific (not shared)

## 📖 Usage

### Quick Create Pass

1. Open an email from someone you want to create a pass for
2. Click **Quick Create Pass** button in the ribbon
3. The sender's name and email are auto-filled
4. Enter the phone number (without country code)
5. Select the country code
6. Click **Create Pass**

### Advanced Pass Creation

1. Open an email
2. Click **Advanced Pass** button
3. Fill in additional details:
   - Personal info (ID number, etc.)
   - Pass details (code, transaction ID, allowed entrances)
   - Security level
   - Access control (areas, sub-areas, row, seat)
   - Options (transferable, open once)
4. Click **Create Pass**

## 🔐 Security Features

- ✅ **Encrypted Storage**: Credentials stored in Office Roaming Settings
- ✅ **HTTPS Only**: All API calls over secure connection
- ✅ **Token-Based Auth**: Uses client API token (no passwords)
- ✅ **Input Validation**: Phone numbers and emails validated before sending
- ✅ **No Backend Changes**: Zero modifications to PAPI code

## 🐛 Troubleshooting

### Add-in doesn't appear in ribbon

- Restart Outlook
- Check that manifest.xml is valid: `npm run validate`
- Ensure server is running: `npm run serve`

### "Not Configured" warning

- Click the Setup button and enter your PAPI credentials
- Test the connection before saving

### Connection test fails

- Verify API Base URL is correct (no trailing slash)
- Check Client ID and API Token
- Ensure Event ID exists in PAPI
- Check network/firewall settings

### Pass creation fails

- Check phone number format (9-15 digits, no country code)
- Verify you have the right permissions in PAPI
- Check browser console for detailed error messages

### CORS errors

- Ensure your PAPI server allows requests from your add-in domain
- Add the domain to PAPI's CORS whitelist

## 📦 Production Deployment

### Option 1: Internal Web Server

1. Host the add-in files on your internal web server (HTTPS required)
2. Update `manifest.xml` URLs to your server
3. Distribute manifest.xml to users

### Option 2: Azure Static Web Apps

1. Deploy to Azure Static Web Apps
2. Update manifest URLs
3. Use Centralized Deployment in Microsoft 365 Admin Center

### Option 3: SharePoint

1. Upload files to SharePoint document library
2. Enable anonymous access for the folder
3. Update manifest with SharePoint URLs

## 🔄 Updates

To update the add-in:

1. Modify the code files
2. Increment version in `manifest.xml`
3. Users will be prompted to update automatically

## 📝 API Reference

### PAPI Endpoint Used

```
POST /passes
```

### Payload Structure

```json
{
  "system_name": "OUTLOOK",
  "system_type": "api",
  "client_id": "123",
  "token": "your-api-token",
  "send_sms_new_pass": true,
  "passes": [
    {
      "event_id": "456",
      "mobile": "501234567",
      "country_code": "972",
      "name": "John",
      "last_name": "Doe",
      "email": "john@example.com",
      "category": 11
    }
  ]
}
```

### Pass Category

All passes created by this add-in use category `11` (Authorized Pass).

## 🛠️ Development

### File Structure

- **manifest.xml**: Add-in configuration and ribbon buttons
- **src/setup/**: Configuration UI
- **src/taskpane/**: Quick and advanced pass creation UIs
- **src/commands/**: Ribbon button handlers
- **src/shared/utils.js**: Shared utilities (API calls, storage)

### Key Functions

- `getConfig()`: Retrieve stored configuration
- `saveConfig(config)`: Save configuration securely
- `createPass(passData, config)`: Call PAPI API to create pass

### Testing

```powershell
# Validate manifest
npm run validate

# Start dev server and sideload
npm run dev

# Stop debugging
npm run stop
```

## 📞 Support

For issues or questions:

- Check the troubleshooting section above
- Review browser console for error messages
- Contact Yabi Technologies support

## 📄 License

MIT License - © 2025 Yabi Technologies

## 🙏 Acknowledgments

Built with Office.js and designed for seamless PAPI integration.

# שליחת טיוטות Outlook | Outlook Drafts Helper

🚀 **A desktop web application for creating multiple Outlook email drafts efficiently**

---

## 📋 תיאור | Description

**עברית:** יישום web המאפשר ליצור מספר רב של טיוטות דוא"ל ב־Microsoft Outlook בפעולה אחת. התוכנה משלבת ממשק web קל לשימוש עם backend בעל יכולת אוטומציה של Outlook.

**English:** A web application that allows you to create multiple email drafts in Microsoft Outlook in a single action. It combines an easy-to-use web interface with a backend capable of automating Outlook.

---

## ✨ תכונות | Features

✅ **יצירת טיוטות מרובות** | Create multiple drafts at once  
✅ **תמיכה במנמענים מרובים** | Support for multiple recipients  
✅ **צירוף קבצים** | File attachment support  
✅ **ממשק עברי** | Full Hebrew UI with RTL support  
✅ **פתוח טיוטה אוטומטי** | Auto-open drafts in Outlook for review  
✅ **ללא שליחה אוטומטית** | Safe - drafts require manual confirmation before sending  

---

## 🏗️ Architecture | ארכיטקטורה

```
outlook/
├── LocalOutlookHelper/        # Backend (.NET 8 Web API)
│   ├── Program.cs             # API endpoints & CORS config
│   ├── OutlookHelper.cs       # Outlook COM interop logic
│   ├── LocalOutlookHelper.csproj
│   ├── Properties/
│   │   └── launchSettings.json # Port 5001 configuration
│   └── appsettings.json
│
└── web/                        # Frontend (HTML/CSS/JavaScript)
    ├── index.html             # Form UI (Hebrew RTL)
    ├── script.js              # Form submission & API calls
    └── style.css              # Styling with Outlook theme
```

### Two-Tier Architecture
- **Frontend**: Lightweight HTML/CSS/JavaScript interface
- **Backend**: ASP.NET Core 8 Web API with Outlook COM interop

---

## 🛠️ Tech Stack | טכנולוגיות

### Backend
- **Framework**: ASP.NET Core 8.0
- **Language**: C#
- **COM Interop**: Microsoft Outlook Interop (v15.0.4797.1004)
- **API Style**: RESTful minimal APIs

### Frontend
- **HTML5** with Hebrew RTL layout
- **CSS3** with Outlook blue theme (#0078d7)
- **JavaScript (ES6)** with FormData API

### Requirements
- **Windows OS** (COM Interop requires Windows)
- **Microsoft Outlook** installed locally
- **Visual Studio 2022** (for development)
- **.NET 8 SDK**

---

## 📦 Installation | התקנה

### Prerequisites / דרישות מוקדמות
1. **Windows 10/11**
2. **Microsoft Outlook** (2016 or later)
3. **.NET 8 SDK** - Download from [dotnet.microsoft.com](https://dotnet.microsoft.com/download)

### Clone & Setup

```bash
# Clone repository
git clone <repository-url>
cd outlook

# Build the solution
dotnet build

# Restore packages (if needed)
dotnet restore
```

---

## ▶️ Running the Application | הפעלת התוכנה

### 1. Start Backend Server
```bash
cd LocalOutlookHelper
dotnet run
```
✅ Backend will run on `http://localhost:5001`

### 2. Open Frontend in Browser
```bash
# Option 1: Open file directly
start web/index.html

# Option 2: Use a local web server (recommended)
# Install Python: python -m http.server 8000 (from web/ directory)
# Or use any local server, then visit: http://localhost:8000
```

---

## 🎯 Usage Guide | מדריך השימוש

### Creating Email Drafts

1. **Fill in Recipients** | הכנסו כתובות דוא"ל
   - Enter multiple email addresses separated by semicolons (`;`)
   - Example: `email1@example.com; email2@example.com`

2. **Add Subject** | הוסיפו נושא
   - Type the email subject line

3. **Write Message** | כתבו הודעה
   - Enter the email body/content

4. **Optional: Attach File** | אופציונלי: צרפו קובץ
   - Click "קובץ מצורף" to select a file to attach

5. **Click Send** | לחצו שלח
   - The backend will create individual Outlook drafts
   - Each draft will open in Outlook for review
   - Success message shows number of drafts created

---

## 🔌 API Endpoints | נקודות ה־API

### POST `/create-drafts`
Creates email drafts based on form data.

**Request Format (multipart/form-data):**
```
- to: "email1@example.com; email2@example.com"
- subject: "Email Subject"
- body: "Email Body Content"
- cv: [File] (optional)
```

**Response (JSON):**
```json
{
  "success": true,
  "message": "3 drafts created successfully",
  "draftCount": 3
}
```

**Error Response:**
```json
{
  "success": false,
  "message": "Error message in Hebrew or English"
}
```

---

## 💻 Code Structure | מבנה הקוד

### Key Files

#### `Program.cs` - API Server
- Configures ASP.NET Core
- Enables CORS for frontend communication
- Defines `/create-drafts` POST endpoint
- Handles file uploads and routing

#### `OutlookHelper.cs` - Outlook Integration
- Contains `CreateDraft()` method
- Uses COM interop to access Outlook.Application
- Creates mail items with recipients, subject, body
- Handles file attachments
- Opens drafts in Outlook without sending

#### `web/script.js` - Frontend Logic
- Form submission handler
- Sends POST request to backend
- Processes FormData with files
- Displays success/error messages

#### `web/style.css` - UI Styling
- RTL (Right-to-Left) layout for Hebrew
- Outlook blue color scheme
- Responsive design
- Custom file upload styling

---

## ⚠️ Important Notes | הערות חשובות

1. **Windows Only**: This application requires Windows and Microsoft Outlook installed
2. **COM Interop**: Outlook must be running or available on the system
3. **Security**: Backend includes CORS but should be restricted in production
4. **Manual Confirmation**: Drafts open in Outlook - user must click Send manually
5. **File Size**: Consider network/attachment size limits

---

## 🐛 Troubleshooting | פתרון בעיות

| Problem | Solution |
|---------|----------|
| Backend won't start | Check if port 5001 is available. Kill process: `netstat -ano \| findstr :5001` |
| "Outlook not found" error | Ensure Microsoft Outlook is installed and accessible |
| CORS error from frontend | Verify backend is running on `http://localhost:5001` |
| File upload fails | Check file size limits and temp directory permissions |
| Drafts not appearing | Ensure Outlook is running and not in offline mode |

---

## 📝 Development

### Modify Backend
```bash
cd LocalOutlookHelper
# Edit OutlookHelper.cs or Program.cs
dotnet run
```

### Modify Frontend
```bash
# Edit web/index.html, script.js, or style.css
# Refresh browser to see changes
```

### Build Release
```bash
dotnet build -c Release
```

---

## 📄 Project Files

- **outlook.sln** - Visual Studio solution file
- **LocalOutlookHelper.csproj** - C# project file with dependencies
- **appsettings.json** - Logging configuration
- **launchSettings.json** - Startup configuration (port 5001)

---

## 🤝 Contributing

This is a personal project. Feel free to fork and customize for your needs!

---

## 📞 Support

For issues or questions:
1. Check the Troubleshooting section
2. Verify all prerequisites are installed
3. Check event logs for Outlook COM errors

---

## 📄 License

This project is provided as-is. Modify for personal or internal use.

---

## 🎉 Version History

**v1.0** - Initial release with core functionality:
- Multiple recipient support
- File attachment handling
- Hebrew UI with RTL layout
- Outlook draft automation via COM interop

---

**עברית | English** - This project is bilingual and supports both languages in the UI.

Made with ❤️ for efficient email draft creation in Microsoft Outlook.

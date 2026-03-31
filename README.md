# 📅 Team Calendar

> **[🇯🇵 日本語版 README はこちら](docs/README.ja.md)**

A Windows desktop application that retrieves meeting information from Outlook calendars, displays response statuses at a glance, and exports to Excel.  
Fetch not only your own calendar but also shared calendars of other team members in one view.

![.NET 10](https://img.shields.io/badge/.NET-10.0-purple)
![Windows Forms](https://img.shields.io/badge/UI-Windows%20Forms-blue)
![License: MIT](https://img.shields.io/badge/License-MIT-green)

---

## ✨ Features

| Feature | Description |
|---------|-------------|
| **Calendar Retrieval** | Reads calendars via Outlook COM (supports recurring appointments) |
| **Multi-User Support** | Fetches your calendar plus shared calendars of other users |
| **Status Detection** | Automatically classifies: Accepted / Tentative / Declined / Organizer / Not Responded |
| **Color-Coded Rows** | Row colors by status (Accepted = green, Tentative = yellow, Declined = red, Organizer = blue) |
| **Summary Cards** | See counts for total, accepted, tentative, and declined at a glance |
| **Excel Export** | Export only accepted meetings to `.xlsx` (powered by ClosedXML) |
| **Debug Log** | Toggleable real-time log panel for troubleshooting |

---

## 📸 Screenshot

```
┌──────────────────────────────────────────────────────────┐
│  📅  Team Calendar                                        │
├──────────────────────────────────────────────────────────┤
│  Period [2025/03/24] ~ [2025/03/28]  ▶ Load  📊 Export    │
│  👥 Users [user1@example.com; user2@...]  ☑ Include self   │
├──────────────────────────────────────────────────────────┤
│  📋 All   │  ✅ Accepted  │  ⏳ Tentative  │  ❌ Declined │
│     12    │      8       │      3        │     1       │
├──────────────────────────────────────────────────────────┤
│  User     │ Subject  │ Start       │ End         │ Status   │
│  Me       │ Standup  │ 03/24 10:00 │ 03/24 11:00 │ Accepted │
│  user1@.. │ 1-on-1   │ 03/24 14:00 │ 03/24 14:30 │ Tentative│
└──────────────────────────────────────────────────────────┘
```

---

## 🔧 Requirements

- **OS**: Windows 10 / 11
- **Runtime**: [.NET 10 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/10.0)
- **Outlook**: Microsoft Outlook (desktop version) must be installed
- **Shared Calendars**: To retrieve other users' calendars, the target user must have granted calendar sharing permissions

---

## 🚀 Getting Started

### Build & Run

```bash
git clone https://github.com/freyWylfred/TeamCalendar.git
cd TeamCalendar
dotnet run --project TeamCalendar
```

### Release Build

```bash
dotnet publish TeamCalendar -c Release -o ./publish
```

Run `TeamCalendar.exe` from the `./publish` folder.

---

## 📖 Usage

### 1. Retrieve Appointments

1. Select the **date range** (defaults to this week, Monday–Friday)
2. To include other users, enter their email addresses separated by semicolons in the **"👥 Target Users"** field
3. Click the **"▶ Load"** button

### 2. Export to Excel

1. After loading appointments, click the **"📊 Excel Export (Accepted)"** button
2. Choose a save location — only accepted (Accepted + Organizer) meetings are exported to `.xlsx`

### 3. Debug Log

- Toggle the **"🔍 Debug Log"** checkbox to show a real-time log panel at the bottom
- Useful for checking Outlook communication status and error details

---

## 🏗 Tech Stack

| Technology | Purpose |
|------------|---------|
| **.NET 10** (Windows Forms) | UI framework |
| **Outlook COM Interop** (`dynamic`) | Access to Outlook calendars |
| **ClosedXML** | Excel (.xlsx) file export |

---

## 📁 Project Structure

```
TeamCalendar/
├── TeamCalendar.slnx           # Solution file
├── .gitignore
├── LICENSE
├── README.md                   # English (this file)
├── docs/
│   └── README.ja.md            # Japanese
└── TeamCalendar/
    ├── TeamCalendar.csproj     # Project definition (.NET 10)
    ├── Program.cs              # Entry point
    ├── Form1.cs                # Main form (logic)
    ├── Form1.Designer.cs       # Main form (UI definition)
    └── Form1.resx              # Resource file
```

---

## ⚠️ Notes

- The **desktop version** of Outlook is required (Outlook on the web / new Outlook are not supported)
- Retrieving other users' calendars requires **calendar sharing permissions** in an Exchange / Microsoft 365 environment
- The "Status" shown for other users reflects **that user's own response status**

---

## 🤝 Contributing

Issues and Pull Requests are welcome!

1. Fork this repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

---

## 📄 License

Released under the [MIT License](LICENSE).

# Simple Scheduler

A clean, privacy‑focused project and task manager that works **fully offline** – no cloud, no accounts, no tracking.  
Organise your work or personal life with **projects → categories → tasks** and track progress with visual percentages.

![Screenshot placeholder](screenshot.png)

## ✨ Features

- **Projects** – create multiple projects, each with its own emoji and title.
- **Categories** – inside every project, add unlimited categories (e.g. “Strategy”, “Tech”, “Marketing”).
- **Tasks** – each task can be marked done/undone, deleted, or added on the fly.
- **Progress bars** – per category and per project, with colour coding (green >60%, purple >30%, orange <30%).
- **100% local** – all data stays in your browser’s `localStorage`. Nothing is sent anywhere.
- **Desktop file sync** (Chrome/Edge) – open or create a `.json` file and changes auto‑save to that file (File System Access API).
- **Portable backups** – export a single project as Markdown (`.md`) or all projects as JSON (full backup).
- **Import from Markdown** – paste Markdown text or load a `.md` file to replace the current project.
- **Responsive** – works on desktop, tablet, and mobile (touch‑friendly).
- **No build step** – single HTML file, open and use immediately.

## 🚀 Quick start

1. Download `launch-scheduler.html`
2. Double‑click to open in your browser (Chrome, Edge, Firefox, Safari)
3. Start creating projects, categories, and tasks

No installation, no server, no hidden costs.

## 🧭 How to use

| Action | How |
|--------|-----|
| Create a project | Click `+ project`, choose an emoji and a name |
| Switch project | Click on any project tab |
| Rename project | Click on the project title |
| Add a category | Click `+ new category` |
| Add a task | Click `+ add task` inside any category |
| Mark task done | Click the checkbox next to the task |
| Reset all tasks in a category | Click `↺ reset` inside the category card |
| Delete a category or task | Hover (desktop) or long‑press (mobile) and click the `×` button |
| Copy current project as Markdown | Click `Copy project` (copies to clipboard) |
| Paste Markdown into current project | Click `Insert project` and paste Markdown into the textarea |
| Save current project as `.md` file | Click `Save project` |
| Load a `.md` file into current project | Click `Load project` and select a `.md` file |
| Full backup (all projects) | Click `Save backup` – downloads a `.json` file |
| Restore from backup | Click `Import backup` and select a previously saved `.json` file |
| Reset everything | Click `reset` (confirmation required) |

### Desktop‑only (Chrome/Edge)

- `open file` – select a `.json` file; the app will use it for auto‑save
- `create file` – create a new `.json` file and save changes directly to it
- `disconnect file` – stop using the file and fall back to `localStorage`

## 🛠️ Technology

- **React 18** (UMD, no build tools)
- **Babel standalone** (JSX in the browser)
- **File System Access API** (optional, desktop Chrome/Edge)
- **localStorage** – persistent storage on all browsers
- **Markdown parser** – import/export using a simple checklist format

## 📁 File formats

### Markdown (`.md`) example
```markdown
# My Project

## Category 1
- [ ] Task one
- [x] Task two (already done)

## Category 2
- [ ] Another task

The app understands - [ ] (incomplete) and - [x] / - [X] (complete).

JSON backup format
Full state: { activeProjectId, projects: [...] }.
Every project contains id, title, emoji, and an array of categories with name, emoji, tasks.

🔒 Privacy
Zero external requests – no analytics, no tracking, no CDN calls (except the initial Google Fonts and React/Babel CDNs if you load the page online).

All your data stays on your device.

You can disconnect from the internet after loading the page and it will still work perfectly.

📱 Mobile support
On Android and iOS, the File System Access buttons (open file / create file) are hidden because the API is not supported.

However, all backup/restore features (.md and .json via file picker) work on mobile browsers.

You can also install the page as a PWA (Progressive Web App) – add a manifest and service worker to enable offline installation.

🧪 Running locally
Just serve the .html file with any static server:

bash
npx serve .
# or
python -m http.server 8000
Or simply open the file with file:// – it works, but some advanced features (like file picker) require a secure context (HTTPS or localhost). For full desktop file support, serve via http://localhost.

🐛 Known limitations
File System Access API is only supported in desktop Chromium browsers (Chrome, Edge, Opera). Firefox and Safari do not implement it.

On iOS (Safari), the file picker for .md / .json works, but you cannot auto‑sync a single file.

The Markdown parser is basic: it expects ## Category headers and - [ ] / - [x] task lines. Other Markdown elements are ignored.

🤝 Contributing
This is a single‑file project. Feel free to fork and modify. If you find a bug or want a feature, open an issue or submit a pull request.

📄 License
MIT – use it for anything, private or commercial.

Made with 💻 and ☕ – because task managers should be simple and yours.

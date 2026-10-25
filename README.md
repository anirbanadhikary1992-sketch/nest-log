# The Nest Log

Bindu & Anirban's nightly check-in. One page each, every evening; a nest that grows one twig per page; ground rules that grow from the log.

- **App:** `index.html`, served by GitHub Pages. No build step.
- **Data:** a Google Sheet ("Nest Log data") in Anirban's Drive, reached through the Apps Script web app in `apps-script/Code.gs`. The app opens only with the shared passcode; owner controls need Anirban's PIN. No personal data lives in this repo.
- **Nightly (Claude scheduled tasks, Anirban's plan):** reads screenshots into pages, writes rule suggestions, drafts app updates from feedback into `preview/` for approval, and sends 9:15 pm reminders. All of it stops when the Control Center switch is off or the matching pause switch in the app is on.

To change the backend address, edit `API_URL` at the top of the script in `index.html`.

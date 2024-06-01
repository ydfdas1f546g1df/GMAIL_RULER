# Getting Started
https://script.google.com

Login to google and create a new project
<br>Go To Project Settings Copy the Script ID and put it into the [.clasp.json](.clasp.json) file

then run the following commands
```bash
npm install -g @google/clasp 
npm install typescript --save-dev
npm install @types/google-apps-script --save-dev
clasp login
clasp push
```
### Now you have a working GMAIL RULER
# Commands
Login to your Google account
```bash
clasp login
```

Clone Project
```bash
clasp clone <script_id>
```

Create a new project
```bash
clasp create --type <docs|sheets|slides|webapp|api>
```

Push your changes
```bash
clasp push
```

Pull the latest changes
```bash
clasp pull
```

Open the script in the browser
```bash
clasp open
```

List all files in the project
```bash
clasp list
```

Create a new file
```bash
clasp create --type <server_js|html|json|css>
```

Push a new version
```bash
clasp push --version description
```

Logout
```bash
clasp logout
```

# Windows Setup Guide - MECE Topic Generator

Complete step-by-step guide to set up the Topic Generator application on a Windows computer.

## 📋 Prerequisites

Before you begin, make sure you have:

1. **Windows 10 or later** (Windows 11 recommended)
2. **Python 3.9 or higher** installed
3. **Internet connection** for downloading dependencies
4. **Azure OpenAI API credentials** (API key and endpoint URL)

---

## 🔧 Step 1: Install Python

### Check if Python is already installed:

1. Press `Win + R` to open Run dialog
2. Type `cmd` and press Enter
3. In the command prompt, type:
   ```cmd
   python --version
   ```

### If Python is NOT installed:

1. Go to [python.org/downloads](https://www.python.org/downloads/)
2. Download the latest Python 3.x version (3.9 or higher)
3. **IMPORTANT**: During installation, check the box **"Add Python to PATH"**
4. Click "Install Now"
5. Wait for installation to complete
6. Restart your computer if prompted

### Verify installation:

Open a new command prompt and run:
```cmd
python --version
```
You should see something like `Python 3.12.x`

---

## 📥 Step 2: Download the Application

### Option A: Download from GitHub (Recommended)

1. Go to: https://github.com/gavrielhan/open-question-part2
2. Click the green **"Code"** button
3. Select **"Download ZIP"**
4. Extract the ZIP file to a location like:
   - `C:\Users\YourName\Desktop\open-question-part2`
   - Or `C:\Programs\open-question-part2`

### Option B: Clone with Git (if you have Git installed)

1. Open Command Prompt or PowerShell
2. Navigate to where you want the project:
   ```cmd
   cd C:\Users\YourName\Desktop
   ```
3. Clone the repository:
   ```cmd
   git clone https://github.com/gavrielhan/open-question-part2.git
   ```

---

## 🗂️ Step 3: Navigate to the Project Folder

1. Open File Explorer
2. Navigate to the folder where you extracted/downloaded the project
3. You should see files like:
   - `topic_generator_app.py`
   - `requirements.txt`
   - `launch_app.bat`
   - `README.md`
   - etc.

---

## 🐍 Step 4: Set Up the Application

### Option A: Automated Setup (Recommended)

1. **Double-click `setup_windows.bat`** in the project folder
   - If Windows shows a security warning, click "More info" → "Run anyway"
   
2. **The script will automatically:**
   - Check Python installation
   - Create virtual environment
   - Install all dependencies
   - Create `.env` file from template

3. **Wait for completion** (2-5 minutes)

### Option B: Manual Setup

1. **Open Command Prompt** in the project folder:
   - In File Explorer, click in the address bar
   - Type `cmd` and press Enter
   - OR right-click in the folder → "Open in Terminal" (Windows 11)
   - OR right-click in the folder → "Open PowerShell window here" (Windows 10)

2. **Create virtual environment:**
   ```cmd
   python -m venv venv
   ```
   This will create a `venv` folder (may take a minute)

3. **Activate the virtual environment:**
   ```cmd
   venv\Scripts\activate
   ```
   You should see `(venv)` at the beginning of your command prompt

4. **Install dependencies:**
   ```cmd
   pip install -r requirements.txt
   ```
   
   This will install:
   - Flask (web framework)
   - pandas (Excel/CSV processing)
   - openpyxl (Excel file support)
   - requests (API calls)
   - python-dotenv (environment variables)
   - PyYAML (YAML parsing)

   **Note:** This may take 2-5 minutes depending on your internet speed.

---

## ⚙️ Step 5: Configure Environment Variables

1. **Locate the example file:**
   - In the project folder, find `env_example.txt`

2. **Create the .env file:**
   - Copy `env_example.txt` and rename it to `.env`
   - **Important:** The file must be named exactly `.env` (with the dot at the beginning)
   - If Windows doesn't let you create a file starting with a dot:
     - Open Notepad
     - Save As → File name: `.env` (with quotes: `".env"`)
     - Save in the project folder

3. **Edit the .env file:**
   - Right-click `.env` → Open with → Notepad
   - Fill in your Azure OpenAI credentials:

   ```env
   OPENAI_API_KEY=your_actual_api_key_here
   OPENAI_API_BASE_URL=https://your-resource.openai.azure.com
   MODEL=gpt-5.1
   AZURE_API_VERSION=2025-04-01-preview
   ```

   **Replace:**
   - `your_actual_api_key_here` with your actual Azure OpenAI API key
   - `https://your-resource.openai.azure.com` with your actual Azure endpoint URL

4. **Save the file** (Ctrl+S)

---

## 🚀 Step 6: Test the Application

1. **Make sure virtual environment is activated:**
   - If you see `(venv)` at the start of your command prompt, you're good
   - If not, run: `venv\Scripts\activate`

2. **Start the application:**
   ```cmd
   python topic_generator_app.py
   ```

3. **What should happen:**
   - A command window will show startup messages
   - Your default browser should open automatically
   - The app will be available at: `http://127.0.0.1:5001`

4. **If you see errors:**
   - Check that the `.env` file exists and has correct values
   - Make sure all dependencies were installed (Step 5)
   - Verify Python version is 3.9 or higher

---

## 🖥️ Step 7: Create Desktop Shortcut (Optional but Recommended)

### Method 1: Using the Provided Script (Easiest)

1. **Double-click** `create_windows_shortcut.bat`
   - If Windows shows a security warning, click "More info" → "Run anyway"

2. **A shortcut will be created** on your Desktop named "Topic Generator"

3. **To use:** Just double-click the shortcut anytime!

### Method 2: Manual Creation

1. Right-click on `launch_app.bat` → "Create shortcut"
2. Cut the shortcut (Ctrl+X)
3. Go to Desktop (Win+D)
4. Paste the shortcut (Ctrl+V)
5. Right-click the shortcut → Properties
6. Click "Change Icon" → Browse → Select `assets\app_icon.ico`
7. Click OK

---

## 📝 Step 8: Using the Application

### Starting the App:

**Option A: Using the shortcut**
- Double-click "Topic Generator" on your Desktop

**Option B: Using the batch file**
- Double-click `launch_app.bat` in the project folder

**Option C: Manual start**
- Open Command Prompt in project folder
- Run: `venv\Scripts\activate`
- Run: `python topic_generator_app.py`

### Using the Web Interface:

1. **Upload a file:**
   - Drag and drop an Excel (.xlsx, .xls) or CSV file
   - Or click to browse and select a file

2. **Select sheet** (if Excel has multiple sheets)

3. **Configure settings:**
   - Enter the column letter containing answers (e.g., "H")
   - Adjust max topics slider (2-15)

4. **Generate topics:**
   - Click "יצירת נושאים אוטומטית" (Generate Topics Automatically)
   - Wait for processing (this may take a few minutes)

5. **Classify responses:**
   - After topics are generated, click "סיווג כל התשובות" (Classify All Responses)
   - The classified file will be saved to your Desktop

---

## 🔧 Troubleshooting

### Problem: "Python is not recognized"

**Solution:**
- Python is not in your PATH
- Reinstall Python and make sure to check "Add Python to PATH"
- Or add Python manually to PATH (advanced)

### Problem: "No module named 'flask'"

**Solution:**
- Virtual environment is not activated
- Run: `venv\Scripts\activate` first
- Then install: `pip install -r requirements.txt`

### Problem: "Port 5001 is already in use"

**Solution:**
- Another instance of the app is running
- Close all command windows
- Or restart your computer

### Problem: "ERROR: .env file not found!"

**Solution:**
- Create `.env` file from `env_example.txt`
- Make sure it's in the same folder as `topic_generator_app.py`
- Make sure it's named exactly `.env` (not `env.txt` or `.env.txt`)

### Problem: "API Error" or "Missing OPENAI_API_KEY"

**Solution:**
- Check your `.env` file has correct values
- Make sure there are no extra spaces
- Verify your API key is valid
- Check your Azure endpoint URL is correct

### Problem: Browser doesn't open automatically

**Solution:**
- Manually open browser
- Go to: `http://127.0.0.1:5001`
- Or: `http://localhost:5001`

### Problem: Can't create .env file (starts with dot)

**Solution:**
- Open Notepad
- Save As → Type filename as: `".env"` (with quotes)
- Or use Command Prompt:
  ```cmd
  copy env_example.txt .env
  ```

---

## 🔄 Updating the Application

If you need to update to the latest version:

1. **Backup your `.env` file** (copy it somewhere safe)

2. **Download the latest version** from GitHub

3. **Replace all files** except:
   - `venv` folder (you can keep it or recreate)
   - `.env` file (restore your backup)

4. **Reinstall dependencies** (if needed):
   ```cmd
   venv\Scripts\activate
   pip install -r requirements.txt --upgrade
   ```

---

## 📞 Getting Help

If you encounter issues:

1. Check this guide's Troubleshooting section
2. Check the main README.md file
3. Verify all prerequisites are met
4. Check that your `.env` file is configured correctly

---

## ✅ Quick Start Checklist

- [ ] Python 3.9+ installed
- [ ] Project downloaded/extracted
- [ ] Virtual environment created (`python -m venv venv`)
- [ ] Virtual environment activated (`venv\Scripts\activate`)
- [ ] Dependencies installed (`pip install -r requirements.txt`)
- [ ] `.env` file created and configured
- [ ] Application tested (`python topic_generator_app.py`)
- [ ] Desktop shortcut created (optional)

---

## 🎉 You're All Set!

Once you've completed all steps, you can:
- Launch the app using the desktop shortcut
- Upload Excel/CSV files
- Generate MECE topics automatically
- Classify responses

Enjoy using the Topic Generator! 🚀


#!/bin/bash
# Script to create a macOS .app bundle for the Topic Generator Tool
# This creates a double-clickable application icon for your desktop

APP_NAME="Topic Generator"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
DESKTOP_DIR="$HOME/Desktop"
APP_PATH="$DESKTOP_DIR/$APP_NAME.app"

echo "Creating macOS application: $APP_NAME.app"
echo "Location: $DESKTOP_DIR"
echo ""

# Create .app bundle structure
mkdir -p "$APP_PATH/Contents/MacOS"
mkdir -p "$APP_PATH/Contents/Resources"

# Create Info.plist
cat > "$APP_PATH/Contents/Info.plist" << 'EOF'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>launcher</string>
    <key>CFBundleIconFile</key>
    <string>AppIcon</string>
    <key>CFBundleIdentifier</key>
    <string>com.samuelneaman.topicgenerator</string>
    <key>CFBundleName</key>
    <string>Topic Generator</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleShortVersionString</key>
    <string>1.0</string>
    <key>CFBundleVersion</key>
    <string>1</string>
    <key>LSMinimumSystemVersion</key>
    <string>10.13</string>
    <key>NSHighResolutionCapable</key>
    <true/>
</dict>
</plist>
EOF

# Create launcher script with proper path injection
cat > "$APP_PATH/Contents/MacOS/launcher" << 'LAUNCHER_EOF'
#!/bin/bash
# Get the directory where this script's project is located
SCRIPT_DIR="__SCRIPT_DIR_PLACEHOLDER__"

# Navigate to script directory
cd "$SCRIPT_DIR" || {
    osascript -e 'display dialog "Error: Could not find application directory!" buttons {"OK"} default button "OK" with icon stop'
    exit 1
}

# Activate virtual environment if it exists
if [ -d "venv" ]; then
    source venv/bin/activate
elif [ -d ".venv" ]; then
    source .venv/bin/activate
fi

# Check if .env file exists
if [ ! -f ".env" ]; then
    osascript -e "display dialog \"Error: .env file not found!\\n\\nPlease create a .env file with your API configuration in:\\n$SCRIPT_DIR\" buttons {\"OK\"} default button \"OK\" with icon stop"
    exit 1
fi

# Open Terminal window and run the web app
osascript <<APPLESCRIPT_EOF
tell application "Terminal"
    activate
    set currentTab to do script "cd \"__SCRIPT_DIR_PLACEHOLDER__\" && echo '🚀 Starting Topic Generator Web App...' && echo '' && echo 'Stopping any existing instances...' && lsof -ti:5001 | xargs kill -9 2>/dev/null || true && sleep 1 && if [ -d \"venv\" ]; then source venv/bin/activate; fi && python3 topic_generator_app.py"
end tell
APPLESCRIPT_EOF
LAUNCHER_EOF

# Inject the actual script directory path
sed -i '' "s|__SCRIPT_DIR_PLACEHOLDER__|$SCRIPT_DIR|g" "$APP_PATH/Contents/MacOS/launcher"

chmod +x "$APP_PATH/Contents/MacOS/launcher"

# Copy bundled icon if available
ICON_SOURCE="$SCRIPT_DIR/assets/AppIcon.icns"
if [ -f "$ICON_SOURCE" ]; then
    cp "$ICON_SOURCE" "$APP_PATH/Contents/Resources/AppIcon.icns"
    echo "🖼️  Attached custom macOS icon from $ICON_SOURCE"
else
    echo "⚠️  AppIcon.icns not found in assets/. Using default icon."
fi

# Create a README for icon customization
cat > "$APP_PATH/Contents/Resources/README_ICON.txt" << 'ICON_README'
This app already ships with a custom icon (assets/AppIcon.icns).
If you ever want to replace it manually:

1. Create an icon file (PNG or ICNS format)
2. Right-click on 'Topic Generator.app' on your Desktop
3. Select 'Get Info'
4. Drag your icon file onto the small icon in the top-left corner of the Info window
5. The custom icon will be applied immediately
ICON_README

echo ""
echo "✅ Application created successfully!"
echo ""
echo "📱 Location: $APP_PATH"
echo ""
echo "To use the app:"
echo "  1. Double-click 'Topic Generator.app' on your Desktop"
echo "  2. If macOS shows a security warning, go to System Preferences → Security & Privacy"
echo "  3. Click 'Open Anyway' to allow the app to run"
echo ""
echo "Note: The app will use the default Terminal icon."
echo "      See $APP_PATH/Contents/Resources/README_ICON.txt for instructions on customizing the icon."
echo ""
echo "To uninstall: Simply drag 'Topic Generator.app' to the Trash"
echo ""


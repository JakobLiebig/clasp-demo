#!/bin/bash

# Quick setup script for clasp demo
# Run this after cloning the repo

echo "🚀 Setting up Data Analyzer Pro..."

# Check if clasp is installed
if ! command -v clasp &> /dev/null; then
    echo "❌ clasp is not installed. Install it with: npm install -g @google/clasp"
    exit 1
fi

echo "✅ clasp found!"

# Login to clasp if not already logged in
echo "📝 Checking clasp login..."
if ! clasp login --status &> /dev/null; then
    echo "🔐 Please login to Google..."
    clasp login
fi

# Create the Apps Script project
echo "📦 Creating Google Sheets project..."
clasp create --type sheets --title "Data Analyzer Pro"

# Push the code
echo "⬆️  Pushing code to Google..."
clasp push

echo ""
echo "✨ Setup complete!"
echo ""
echo "📋 Next steps:"
echo "1. Open Google Sheets: https://sheets.google.com"
echo "2. Create a new spreadsheet"
echo "3. Add some sample data"
echo "4. Refresh the page"
echo "5. Look for the '📊 Data Analyzer Pro' menu"
echo ""
echo "💡 Tip: Run 'clasp open' to view your project in the Apps Script editor"
echo ""

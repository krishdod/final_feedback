#!/bin/bash
# Build script for Render deployment

echo "🚀 Starting build process..."

# Upgrade pip
pip install --upgrade pip

# Install dependencies
echo "📦 Installing Python dependencies..."
pip install -r requirements.txt

echo "✅ Build completed successfully!"

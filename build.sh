#!/bin/bash
echo "🚀 Starting Render build process..."

# Install system dependencies for WeasyPrint on Render
apt-get update
apt-get install -y \
    python3-dev \
    libcairo2 \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libgdk-pixbuf2.0-0 \
    libffi-dev \
    shared-mime-info \
    libpangoft2-1.0-0

# Install Python dependencies
pip install -r requirements.txt

echo "✅ Build completed successfully!"
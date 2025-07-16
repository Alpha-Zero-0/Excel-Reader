#!/bin/bash

# Excel Reader Tool Startup Script
# This script starts the Streamlit application

echo "Starting Excel Reader Tool..."
echo "Once started, open your browser and go to: http://localhost:8501"
echo ""

# Navigate to the script directory
cd "$(dirname "$0")"

# Run the Streamlit application
/Users/xuan/.pyenv/versions/3.10.12/bin/python -m streamlit run excel_reader.py --server.address localhost --server.port 8501

echo ""
echo "Excel Reader Tool has stopped."

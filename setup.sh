#!/bin/bash

# Create virtual environment with Python 3.10 (which has better compatibility)
python3.10 -m venv venv

# If Python 3.10 is not available, try to install it
if [ $? -ne 0 ]; then
    echo "Python 3.10 not found. Installing Python 3.10..."
    if [[ "$OSTYPE" == "darwin"* ]]; then
        # macOS
        brew install python@3.10
        python3.10 -m venv venv
    else
        # Linux
        sudo apt-get update
        sudo apt-get install python3.10 python3.10-venv
        python3.10 -m venv venv
    fi
fi

# Activate virtual environment
source venv/bin/activate

# Upgrade pip
python -m pip install --upgrade pip

# Install PyTorch and dependencies
echo "Installing PyTorch..."
python -m pip install torch --index-url https://download.pytorch.org/whl/cpu

# Verify PyTorch installation
if ! python -c "import torch; print('PyTorch version:', torch.__version__)" ; then
    echo "PyTorch installation failed!"
    exit 1
fi

# Install transformers and other dependencies
echo "Installing transformers and dependencies..."
python -m pip install transformers
python -m pip install sentencepiece
python -m pip install accelerate
python -m pip install python-docx  # Added support for .docx files
python -m pip install nltk
python -c "import nltk; nltk.download('punkt')"

# Verify all installations
echo "Verifying installations..."
python -c "
import torch
import transformers
from docx import Document
print('PyTorch version:', torch.__version__)
print('Transformers version:', transformers.__version__)
"

echo "Setup complete! To run the script, first activate the virtual environment with:"
echo "source venv/bin/activate"
echo "Then run:"
echo "python doc_rephraser.py" 
# Create virtual environment
python -m venv venv

# Activate virtual environment
.\venv\Scripts\Activate.ps1

# Install requirements
pip install -r requirements.txt

Write-Host "Virtual environment created and dependencies installed!" -ForegroundColor Green

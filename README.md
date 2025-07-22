# 🧪 Python Labs

A collection of Python mini-projects, experiments, and practical exercises for learning, testing, and prototyping.

---

## 🚀 Features

- Modular Python scripts and notebooks
- Isolated per-lab environments using `venv`
- Simple setup and execution
- Ideal for experimentation and skill sharpening

---

## 🛠️ Tech Stack

- Python 3.8+
- Virtual Environments (`venv`)
- Ubuntu / Linux recommended
- Optional: pip, requirements.txt

---

## ▶️ Running a Lab

Use a virtual environment to keep each lab isolated and clean.

```bash
# 1. Update package lists
sudo apt update

# 2. Install venv module (if not installed)
sudo apt install python3-venv

# 3. Create a virtual environment (custom name allowed)
python3 -m venv ~/venvs/excel-env

# 4. Activate the virtual environment
source ~/venvs/excel-env/bin/activate

# 5. Install dependencies if any
# pip install -r requirements.txt

# 6. Run your script
# python lab1/main.py

# 7. Deactivate when done
deactivate

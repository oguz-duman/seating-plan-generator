# Exam Seating Plan Generator

This script generates randomized exam seating plans from an Excel configuration file.

Students are distributed across classrooms using a round-robin strategy to ensure a balanced placement.  

The output consists of two Excel files:
- one containing only student numbers
- one containing student numbers together with student names
inside the folder created using the exam name.

The original classroom layout is preserved.

---

### Installation
```bash
git clone https://github.com/oguz-duman/seating-plan-generator
```

```bash
pip install -r requirements.txt
```

```bash
python main.py
```

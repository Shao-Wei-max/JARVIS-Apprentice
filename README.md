# JARVIS Apprentice v0.7
Python 3.11 recommended. Install: `python -m pip install -r requirements.txt`.
UI: `python -m streamlit run app.py`.
Put the original files (unchanged names) under dataset/ (nested folders allowed).
Smoke test: `python experiment_runner.py --split train --limit 1 --out experiments/smoke`
Validation with approved training memory: `python experiment_runner.py --split validation --memory workspace/memory.db --out experiments/validation`
Only after freezing parameters and training-only memory, run test with --split test and a separate output folder. No automatic learning in runner. Without --memory both conditions have no learned rules. Reference is used only by evaluator. Page alignment is provisional and overlapping; subset PDF page numbers are local. Content integration is not included in batch evaluation. Manual metrics remain null. Do not publish private course files or generated outputs.

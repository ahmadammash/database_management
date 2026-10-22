# Import the required libraries
import ast
import sys
from pathlib import Path

import pandas as pd

pd.set_option("display.max_columns", None)  # Ensure all columns are displayed
pd.set_option("display.max_rows", None)  # Ensure all rows are displayed

# Add the project directory to the system path safely
current_dir = Path(__file__).resolve().parent
project_root = next(
    (
        p
        for p in [current_dir] + list(current_dir.parents)
        if (p / "pyproject.toml").exists()
    ),
    current_dir,
)

if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

code_files = project_root / "code_files"
code_files.mkdir(parents=True, exist_ok=True)  # Ensure directory exists
functions_summary = code_files / "FUNCTIONS_SUMMARY.MD"


def extract_functions(root_path: Path):
    """
    This function extracts the locations, names, input parameters, and descriptions of all user-defined functions and adds it to the FUNCTIONS_SUMMARY.md file.
    """
    records = []

    # Define the counter variable for the number of the function in the md file
    i = 1

    # Walk through all .py files recursively, ignoring virtual environments/caches/git
    for path in root_path.rglob("*.py"):
        if any(
            ignored in path.parts
            for ignored in [
                "__pycache__",
                ".venv",
                "venv",
                ".git",
                "tkinter_course",
                "data_tasks",
            ]
        ):
            continue

        try:
            with open(path, "r", encoding="utf-8") as f:
                tree = ast.parse(f.read(), filename=str(path))

            for node in ast.walk(tree):
                if isinstance(node, ast.FunctionDef) and node.name != "main":
                    # Extract function docstring if present
                    docstring = ast.get_docstring(node) or "No description"
                    first_line = docstring.strip().split("\n")[0]  # First line only

                    # Extract argument names
                    args = [arg.arg for arg in node.args.args]

                    records.append(
                        {
                            "Number": i,
                            "Folder": path.parent.name,
                            "File": path.name,
                            "Function": node.name,
                            "Arguments": ", ".join(args),
                            "Description": first_line,
                        }
                    )

                    # Increment the counter variable
                    i += 1
        except Exception as e:
            print(f"Skipped {path} due to error: {e}")

    return pd.DataFrame(records)


if __name__ == "__main__":
    df = extract_functions(project_root)

    # Generate markdown table and replace non-breaking spaces (\u00a0) with standard spaces
    markdown_content = df.to_markdown(index=False).replace("\u00a0", " ")

    # Save clean Markdown table
    with open(functions_summary, "w", encoding="utf-8") as f:
        f.write("# Master Function Directory\n\n")
        f.write(markdown_content)
        f.write("\n")

import csv
import random
import re
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[1]
SRC_CSV = BASE_DIR / "DepartmentPasswords_CONFIDENTIAL.csv"
OUT_CSV = BASE_DIR / "DepartmentPasswords_CONFIDENTIAL_NEW.csv"
OUT_SECRETS = BASE_DIR / ".streamlit" / "secrets_NEW.toml"


def _read_departments_from_csv(path: Path) -> list[str]:
    if not path.exists():
        raise FileNotFoundError(f"Missing {path.name}")

    with path.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        departments: list[str] = []
        for row in reader:
            dept = (row.get("Department") or "").strip()
            if not dept or dept == "[MASTER]":
                continue
            departments.append(dept)

    if not departments:
        raise RuntimeError(f"No departments found in {path.name}")

    return departments


def _prefix(dept: str) -> str:
    # Simple readable prefix from initials.
    words = re.findall(r"[A-Za-z0-9]+", dept.upper())
    initials = "".join(w[0] for w in words if w)
    if len(initials) >= 3:
        return initials[:4]
    flat = "".join(words)
    return (flat[:4] or "DEPT")


def _toml_key(name: str) -> str:
    # Quote anything that isn't a plain TOML bare key.
    if re.fullmatch(r"[A-Za-z0-9_-]+", name):
        return name
    return '"' + name.replace('\\', '\\\\').replace('"', '\\"') + '"'


def main() -> None:
    departments = _read_departments_from_csv(SRC_CSV)

    used: set[str] = set()

    def unique_pwd(pref: str) -> str:
        # Simple but not guessable: PREFIX-######
        for _ in range(20000):
            pwd = f"{pref}-{random.randint(0, 999999):06d}"
            if pwd not in used:
                used.add(pwd)
                return pwd
        raise RuntimeError("Failed to generate unique password")

    master_pwd = f"MASTER-{random.randint(0, 99999999):08d}"
    used.add(master_pwd)

    dept_pw = {dept: unique_pwd(_prefix(dept)) for dept in departments}

    # Write NEW CSV
    with OUT_CSV.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Department", "Password"])
        w.writerow(["[MASTER]", master_pwd])
        for dept in departments:
            w.writerow([dept, dept_pw[dept]])

    # Write NEW secrets TOML
    OUT_SECRETS.parent.mkdir(parents=True, exist_ok=True)
    lines = [f"MASTER_PASSWORD = \"{master_pwd}\"", "", "[DEPARTMENT_PASSWORDS]"]
    for dept in departments:
        lines.append(f"{_toml_key(dept)} = \"{dept_pw[dept]}\"")
    lines += [
        "",
        "# Optional (only if using Supabase persistence)",
        "SUPABASE_URL = \"\"",
        "SUPABASE_KEY = \"\"",
    ]
    OUT_SECRETS.write_text("\n".join(lines) + "\n", encoding="utf-8")

    print("Generated NEW password files:")
    print(f"- {OUT_CSV}")
    print(f"- {OUT_SECRETS}")
    print("(No passwords printed to console.)")


if __name__ == "__main__":
    main()

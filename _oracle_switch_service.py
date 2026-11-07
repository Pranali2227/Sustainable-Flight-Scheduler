from pathlib import Path

env_file = Path(__file__).with_name(".env")
lines = env_file.read_text(encoding="utf-8").splitlines()
updated = []
for line in lines:
    if line.startswith("ORACLE_DSN="):
        updated.append("ORACLE_DSN=localhost:1521/XE")
    else:
        updated.append(line)
env_file.write_text("\n".join(updated) + "\n", encoding="utf-8")
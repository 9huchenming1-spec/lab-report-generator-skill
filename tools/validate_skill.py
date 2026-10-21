"""Portable metadata/resource validation for this repository; no Codex install needed."""

from pathlib import Path
import re
import sys

import yaml

ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "skills/lab-report-generator"


def validate():
    text = (SKILL / "SKILL.md").read_text(encoding="utf-8")
    match = re.match(r"\A---\n(.*?)\n---\n", text, re.S)
    assert match, "Missing YAML frontmatter"
    metadata = yaml.safe_load(match[1])
    assert metadata["name"] == SKILL.name
    assert 0 < len(metadata["description"]) <= 1024
    assert len(metadata["name"]) <= 64 and re.fullmatch(r"[a-z0-9]+(?:-[a-z0-9]+)*", metadata["name"])
    for target in re.findall(r"\]\((references/[^)]+)\)", text):
        assert (SKILL / target).is_file(), f"Missing resource: {target}"
    ui = yaml.safe_load((SKILL / "agents/openai.yaml").read_text(encoding="utf-8"))
    assert "$lab-report-generator" in ui["interface"]["default_prompt"]
    assert 25 <= len(ui["interface"]["short_description"]) <= 64
    assert ui.get("policy", {}).get("allow_implicit_invocation", True)
    for source in (SKILL / "scripts").glob("*.py"):
        compile(source.read_text(encoding="utf-8"), str(source), "exec")
    print("Skill metadata, resource links, invocation and Python syntax are valid.")


if __name__ == "__main__":
    try:
        validate()
    except (AssertionError, KeyError, OSError, yaml.YAMLError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        sys.exit(1)

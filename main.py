"""Run a publisher-specific methane-decomposition data extractor."""

import argparse
import os
import subprocess
import sys
from pathlib import Path


REPOSITORY_DIR = Path(__file__).resolve().parent
EXTRACTORS = {
    "elsevier": REPOSITORY_DIR / "extractors" / "elsevier_extractor.py",
    "mdpi": REPOSITORY_DIR / "extractors" / "mdpi_extractor.py",
    "rsc": REPOSITORY_DIR / "extractors" / "rsc_extractor.py",
    "wiley": REPOSITORY_DIR / "extractors" / "wiley_extractor.py",
}


def parse_args() -> tuple[argparse.Namespace, list[str]]:
    parser = argparse.ArgumentParser(
        description="Choose the publisher and model provider, then forward the remaining extractor options."
    )
    parser.add_argument("--publisher", choices=["elsevier", "mdpi", "rsc", "wiley"])
    parser.add_argument("--provider", choices=["openai", "bedrock"], default="openai")
    parser.add_argument("--list", action="store_true", help="List the included publisher/provider combinations.")
    return parser.parse_known_args()


def main() -> None:
    args, extractor_args = parse_args()
    if args.list:
        for publisher in sorted(EXTRACTORS):
            print(f"{publisher}: openai, bedrock")
        return

    if args.publisher is None:
        raise SystemExit("--publisher is required unless --list is used.")

    script = EXTRACTORS[args.publisher]

    if not any(arg == "--prompts-dir" or arg.startswith("--prompts-dir=") for arg in extractor_args):
        extractor_args.extend(["--prompts-dir", str(REPOSITORY_DIR / "prompts")])

    child_env = os.environ.copy()
    child_env["LLM_PROVIDER"] = args.provider
    result = subprocess.run([sys.executable, str(script), *extractor_args], check=False, env=child_env)
    raise SystemExit(result.returncode)


if __name__ == "__main__":
    main()

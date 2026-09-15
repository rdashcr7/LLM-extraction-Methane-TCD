"""Run a publisher-specific methane-decomposition data extractor."""

import argparse
import subprocess
import sys
from pathlib import Path


REPOSITORY_DIR = Path(__file__).resolve().parent
EXTRACTORS = {
    ("elsevier", "openai"): REPOSITORY_DIR / "helper.py",
    ("elsevier", "bedrock"): REPOSITORY_DIR / "extractors" / "elsevier_bedrock.py",
    ("mdpi", "openai"): REPOSITORY_DIR / "extractors" / "mdpi_openai.py",
    ("rsc", "openai"): REPOSITORY_DIR / "extractors" / "rsc_openai.py",
    ("wiley", "openai"): REPOSITORY_DIR / "extractors" / "wiley_openai.py",
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
        for publisher, provider in sorted(EXTRACTORS):
            print(f"{publisher}: {provider}")
        return

    if args.publisher is None:
        raise SystemExit("--publisher is required unless --list is used.")

    script = EXTRACTORS.get((args.publisher, args.provider))
    if script is None:
        available = ", ".join(provider for publisher, provider in EXTRACTORS if publisher == args.publisher)
        raise SystemExit(f"No {args.provider} extractor is included for {args.publisher}. Available: {available}.")

    if not any(arg == "--prompts-dir" or arg.startswith("--prompts-dir=") for arg in extractor_args):
        extractor_args.extend(["--prompts-dir", str(REPOSITORY_DIR / "prompts")])

    result = subprocess.run([sys.executable, str(script), *extractor_args], check=False)
    raise SystemExit(result.returncode)


if __name__ == "__main__":
    main()

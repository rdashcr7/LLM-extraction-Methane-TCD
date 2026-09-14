"""Command-line entry point for the methane-decomposition data extractor."""

import argparse
import os

import helper


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Extract methane-decomposition datapoints from Elsevier XML artifacts "
            "and curated figure images."
        )
    )
    parser.add_argument("--mapping-xlsx", default="Elsevier dois.xlsx", help="Excel workbook with 'Paper' and 'DOIs' columns.")
    parser.add_argument("--xml-root", default="./XMLs", help="Root directory containing XML-derived artifacts for each paper.")
    parser.add_argument("--pdf-base-dir", default="./PDFs", help="Base directory containing PDFs/<Paper>/images/.")
    parser.add_argument("--output-prefix", default="./elsevier_extraction_results", help="Output path prefix, without an extension.")
    parser.add_argument("--paper", action="append", default=None, help="Process only this Paper value. Repeat for multiple papers.")
    parser.add_argument("--fresh-start", action="store_true", help="Remove any saved progress file before starting.")
    parser.add_argument("--prompts-dir", default=os.environ.get("PROMPTS_DIR", "./prompts"), help="Directory containing prompt template text files.")
    parser.add_argument("--resume-from-paper-number", "--resume-from-paper", type=int, default=None, help="Process papers whose numeric label or workbook order is at least this value.")
    parser.add_argument("--send-all-figures-to-llm", "--no-figure-skip", "--disable-figure-skip", action="store_true", help="Send every curated figure image to the LLM.")
    parser.add_argument("--audit-output-dir", default=None, help="Directory for audit reports. Default: <output-prefix>_audit_reports.")
    parser.add_argument("--no-audit-reports", action="store_true", help="Do not generate audit reports.")
    parser.add_argument("--figure-max-output-tokens", type=int, default=helper.FIGURE_MAX_OUTPUT_TOKENS, help="Maximum output tokens for figure extraction calls.")
    parser.add_argument("--figure-schema-mode", choices=["compact", "full"], default=helper.FIGURE_SCHEMA_MODE, help="Use the compact or full figure-row schema.")
    parser.add_argument("--openai-model", default=os.environ.get("OPENAI_MODEL_ID", helper.MODEL_ID), help="OpenAI model ID.")
    parser.add_argument("--openai-base-url", default=os.environ.get("OPENAI_BASE_URL", helper.OPENAI_BASE_URL), help="OpenAI API base URL.")
    parser.add_argument("--openai-reasoning-effort", default=helper.OPENAI_REASONING_EFFORT, help="Reasoning effort: low, medium, high, xhigh, or none.")
    parser.add_argument("--openai-text-verbosity", default=helper.OPENAI_TEXT_VERBOSITY, help="Text verbosity: low, medium, high, or none.")
    parser.add_argument("--llm-call-delay-seconds", type=float, default=helper.LLM_CALL_DELAY_SECONDS, help="Delay before each API call.")
    parser.add_argument("--throttle-max-retries", type=int, default=helper.LLM_THROTTLE_MAX_RETRIES, help="Retries after rate/quota limiting.")
    parser.add_argument("--throttle-base-delay-seconds", type=float, default=helper.LLM_THROTTLE_BASE_DELAY_SECONDS, help="Initial throttling backoff delay.")
    parser.add_argument("--throttle-max-delay-seconds", type=float, default=helper.LLM_THROTTLE_MAX_DELAY_SECONDS, help="Maximum throttling backoff delay.")
    parser.add_argument("--throttle-jitter-seconds", type=float, default=helper.LLM_THROTTLE_JITTER_SECONDS, help="Random jitter added to throttling delays.")
    return parser.parse_args()


def _print_run_configuration() -> None:
    print("\n" + "=" * 70)
    print("Run configuration")
    print("-" * 70)
    print(f"OpenAI model          : {helper.MODEL_ID}")
    print(f"OpenAI base URL       : {helper.OPENAI_BASE_URL}")
    print(f"Figure schema         : {helper.FIGURE_SCHEMA_MODE}")
    print(f"Max figure tokens     : {helper.FIGURE_MAX_OUTPUT_TOKENS}")
    print(f"LLM call delay        : {helper.LLM_CALL_DELAY_SECONDS} s")
    print(f"Throttle retries      : {helper.LLM_THROTTLE_MAX_RETRIES}")
    print("=" * 70 + "\n")


def main() -> None:
    args = parse_args()
    helper._set_prompts_dir(args.prompts_dir)
    helper._set_openai_runtime_settings(
        model_id=args.openai_model,
        base_url=args.openai_base_url,
        reasoning_effort=args.openai_reasoning_effort,
        text_verbosity=args.openai_text_verbosity,
    )
    helper._set_llm_rate_limit_settings(
        call_delay_seconds=args.llm_call_delay_seconds,
        throttle_max_retries=args.throttle_max_retries,
        throttle_base_delay_seconds=args.throttle_base_delay_seconds,
        throttle_max_delay_seconds=args.throttle_max_delay_seconds,
        throttle_jitter_seconds=args.throttle_jitter_seconds,
    )
    helper._set_figure_runtime_settings(
        figure_max_output_tokens=args.figure_max_output_tokens,
        figure_schema_mode=args.figure_schema_mode,
    )

    if helper.openai_client is None:
        raise RuntimeError(
            "OpenAI client is not configured. Set OPENAI_API_KEY; OPENAI_BASE_URL and OPENAI_MODEL_ID are optional."
        )

    _print_run_configuration()
    helper.run_extraction_pipeline(
        mapping_xlsx=args.mapping_xlsx,
        parsed_xml_dir=args.xml_root,
        pdf_base_dir=args.pdf_base_dir,
        output_prefix=args.output_prefix,
        fresh_start=args.fresh_start,
        paper_filters=args.paper,
        resume_from_paper_number=args.resume_from_paper_number,
        send_all_figures_to_llm=args.send_all_figures_to_llm,
        save_audit_reports=not args.no_audit_reports,
        audit_output_root=args.audit_output_dir,
    )


if __name__ == "__main__":
    main()

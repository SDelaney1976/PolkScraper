# src/validator/main.py
import os
import sys
import traceback
import runpy

def run(input_path: str | None = None):
    """
    Execute validator.validate_address in-process, passing an optional input path.

    This avoids launching another Tk root or re-entering your app.
    We pass the file path via both sys.argv and an env var so your
    existing validate_address script can pick it up either way.
    """
    # Capture original argv to restore later
    orig_argv = list(sys.argv)
    try:
        if input_path:
            # Many legacy scripts read argv[1], so provide it.
            sys.argv = [orig_argv[0], input_path]
            # Also expose via env for scripts that prefer env-based input.
            os.environ["VALIDATOR_INPUT"] = input_path
        else:
            # No input provided; present empty argv to avoid accidental recursion.
            sys.argv = [orig_argv[0]]

        # Flag so validate_address can skip any GUI file pickers if it supports it.
        os.environ["VALIDATOR_HEADLESS"] = "1"

        # Run as a module so its existing __main__ path executes.
        runpy.run_module("validator.validate_address", run_name="__main__")

    except SystemExit:
        # Allow scripts that call sys.exit()
        pass
    except Exception:
        traceback.print_exc()
        raise
    finally:
        # Restore argv for safety
        sys.argv = orig_argv
        # Best-effort cleanup of the env hints
        os.environ.pop("VALIDATOR_INPUT", None)
        os.environ.pop("VALIDATOR_HEADLESS", None)

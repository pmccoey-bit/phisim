"""Run the PHISIM school-validation applications in the required order."""

import argparse
from pathlib import Path
import subprocess
import sys


APPLICATIONS = (
    ('Update PHISIM qualification form', 'update_phisim_qualification_form.py'),
    ('Create current evaluation checklist', 'create_current_evaluation_checklist.py'),
)


def main():
    parser = argparse.ArgumentParser(
        description=__doc__,
        epilog='Examples:\n  py run_school_validation.py\n  py run_school_validation.py --dry-run',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Preview both applications without saving changes.',
    )
    args = parser.parse_args()
    folder = Path(__file__).resolve().parent

    for label, filename in APPLICATIONS:
        command = [sys.executable, str(folder / filename)]
        if args.dry_run:
            command.append('--dry-run')
        print(f'\n== {label} ==', flush=True)
        result = subprocess.run(command, cwd=folder, check=False)
        if result.returncode:
            raise SystemExit(result.returncode)


if __name__ == '__main__':
    main()

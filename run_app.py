import sys

if not ((3, 10) <= sys.version_info[:2] <= (3, 11)):
    raise SystemExit("Answer Sheet Studio requires Python 3.10 or 3.11.")

from app.main import run

if __name__ == "__main__":
    run()

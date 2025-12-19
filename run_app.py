import sys

if sys.version_info < (3, 10):
    raise SystemExit("Answer Sheet Studio requires Python 3.10+.")

from app.main import run

if __name__ == "__main__":
    run()

import sys
import warnings
from pathlib import Path

# Put the project root on sys.path so `import processor` etc. work without
# needing an installed package.
sys.path.insert(0, str(Path(__file__).parent))

# We use regex capturing groups in the client-matching rules as a readability
# convenience; pandas doesn't like that for str.contains(). The match is
# boolean-only, so the warning is noise.
warnings.filterwarnings(
    "ignore",
    message=".*pattern is interpreted as a regular expression.*",
    category=UserWarning,
)

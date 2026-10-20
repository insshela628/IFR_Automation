# Tool's own IFRAutomation.archive_approved_in_ifr_client, called on a light host (full IFRAutomation needs root+config)
import sys, logging, types
from pathlib import Path
sys.path.insert(0, r"D:\1.dev\work\L03-pipeline\ifr-automation")
import ifr_automation_v10 as v
logging.basicConfig(level=logging.INFO, format="%(message)s")
P = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\Waterloo")
h = types.SimpleNamespace(dry_run="--apply" not in sys.argv, logger=logging.getLogger("archive"))
h._to_long_path = types.MethodType(v.IFRAutomation._to_long_path, h)
print(v.IFRAutomation.archive_approved_in_ifr_client(h, P))

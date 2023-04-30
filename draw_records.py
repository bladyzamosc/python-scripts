import os
import shutil
from datetime import datetime

from randomizer import Randomizer

# datetime object containing current date and time
now = datetime.now()
print("now =", now)
# dd/mm/YY H:M:S
dt_string = now.strftime("%d%m%Y_%H%M%S")
print("date and time =", dt_string)

files = [f for f in os.listdir(os.path.join(os.curdir, "files")) if not f.startswith("~") and f.endswith("xlsx")]
print(files)

from_path = os.path.join(os.curdir, "files")
to_path = os.path.join(os.curdir, "files", dt_string);
os.makedirs(to_path, exist_ok=True, mode=777)

for f in files:
    file = os.path.join(from_path, f)
    to_file = os.path.join(to_path, f)
    shutil.copyfile(file, to_file)
    r = Randomizer(to_file)
    r.randomize()

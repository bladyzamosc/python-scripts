import os
import shutil
from datetime import datetime

from randomizer import Randomizer

# datetime object containing current date and time
now = datetime.now()
# dd/mm/YY H:M:S
dt_string = now.strftime("%d%m%Y_%H%M%S")
path = os.path.join(os.curdir, "files")
print("Szukanie plikow w folderze: ", os.path.join(os.curdir, "files"))
os.makedirs(path, exist_ok=True, mode=777)
files = [f for f in os.listdir(path) if not f.startswith("~") and f.endswith("xlsx")]
from_path = path
to_path = os.path.join(os.curdir, "files", dt_string);
os.makedirs(to_path, exist_ok=True, mode=777)

for f in files:
    file = os.path.join(from_path, f)
    to_file = os.path.join(to_path, f)
    shutil.copyfile(file, to_file)
    print("Przetwarzanie pliku: ", to_file)
    r = Randomizer(to_file)
    r.randomize()

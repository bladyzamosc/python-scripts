import os

from randomizer import Randomizer

files = [f for f in os.listdir(os.path.join(os.curdir, "files")) if not f.startswith("~")]
print(files)

for f in files:
    r = Randomizer(f, os.path.join(os.curdir, "files"))
    r.randomize()






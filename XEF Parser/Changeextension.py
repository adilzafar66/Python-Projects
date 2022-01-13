from pathlib import Path
import zipfile
from shutil import copyfile
import os

def extractXefFromZef(filePath):
    
    p = Path(filePath)
    print(p.rename(p.with_suffix('.zip')))
##    filePath = p.rename(p.with_suffix('.zip'))
##    unzipDir = os.path.join(os.path.dirname(filePath),Path(filePath).stem)
##    with zipfile.ZipFile(filePath, 'r') as zip_ref:
##        zip_ref.extractall(unzipDir)
##
##    return list(Path(unzipDir).glob('*.xef'))

extractXefFromZef("C:\\XEFParser\\ZEFSample.zef")

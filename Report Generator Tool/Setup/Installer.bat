@ECHO OFF
pip install pipreqs
cd.. 
pipreqs . --force
pip install -r requirements.txt
PAUSE
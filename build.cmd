@echo off
echo Checking required packages...
pip3 install -r requirements.txt
echo Generating exe...
pyinstaller main.py ^
--name mredak-rotationsmanager ^
--specpath ./build --distpath . ^
--onefile --noconsole --clean ^
--exclude-module numpy ^
--add-data ../template.docx:. ^
--add-data ../theme/azure.tcl:theme ^
--add-data ../theme/theme/:theme/theme
echo Done.
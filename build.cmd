pyinstaller main.py ^
--specpath ./build --distpath . ^
--onefile --noconsole --clean ^
--add-data ../LICENSE:. ^
--add-data ../template.docx:. ^
--add-data ../theme/azure.tcl:theme ^
--add-data ../theme/theme/:theme/theme
version = 0.0.15
build-pyinstaller:
	@echo "build using using pyinstaller"
	pyinstaller .\gui.py --add-data "jira.ico;." --add-data ".env;." --add-data ".\TemplatePayload.json;." --add-data ".\help.txt;." --hiddenimport win32timezone --icon .\jira.ico
	make zip
	@echo "to make an installer run make installer, this will generate a setup.exe in the directory Output"
zip:
	powershell Compress-Archive -Path .\dist\ -DestinationPath .\releases\Jira-Bridge-$(version).zip
installer:
	iscc .\setup-generator.iss

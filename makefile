version = 0.1.7
build-pyinstaller:
	@echo "build using using pyinstaller"
	pyinstaller .\gui.py --windowed --noconfirm --name "jira-flow" --add-data "jira.ico;." --add-data ".env;." --add-data ".\TemplatePayload.json;." --add-data ".\help.txt;." --hiddenimport win32timezone --icon .\jira.ico
	make zip
	@echo "to make an installer run make installer, this will generate a setup.exe in the directory Output"
build-pyinstaller-debug:
	@echo "build using using pyinstaller"
	pyinstaller .\gui.py --noconfirm --name "jira-flow" --add-data "jira.ico;." --add-data ".env;." --add-data ".\TemplatePayload.json;." --add-data ".\help.txt;." --hiddenimport win32timezone --icon .\jira.ico
	make zip
	@echo "to make an installer run make installer, this will generate a setup.exe in the directory Output"
zip:
	powershell Compress-Archive -Force -Path .\dist\ -DestinationPath .\releases\Jira-Bridge-$(version).zip
installer: build-pyinstaller
	iscc /dMyAppVersion=$(version) .\setup-generator.iss
installer: build-pyinstaller-debug
	iscc /dMyAppVersion=$(version) .\setup-generator.iss

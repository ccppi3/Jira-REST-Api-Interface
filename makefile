version = $(shell git describe --tags)
build-pyinstaller: refresh-version 
	@echo "build using using pyinstaller"
	pyinstaller .\gui.py --version-file exeVersionInfo.txt --windowed --noconfirm --name "jira-flow" --add-data "jira.ico;." --add-data ".env;." --add-data ".\TemplatePayload.json;." --add-data ".\help.txt;." --add-data ".venv\Lib\site-packages\tkinterdnd2;." --hiddenimport win32timezone --icon .\jira.ico --collect-all tkinterdnd2
	@echo "to make an installer run make installer, this will generate a setup.exe in the directory Output"
	@echo "to make a zip run make zip"
build-pyinstaller-debug: refresh-version
	@echo "build using using pyinstaller"
	pyinstaller .\gui.py --version-file exeVersionInfo.txt --noconfirm --name "jira-flow" --add-data "jira.ico;." --add-data ".env;." --add-data ".\TemplatePayload.json;." --add-data ".\help.txt;." --add-data ".venv\Lib\site-packages\tkinterdnd2;." --hiddenimport win32timezone --icon .\jira.ico --collect-all tkinterdnd2
	@echo "to make an installer run make installer, this will generate a setup.exe in the directory Output"
	@echo "to make a zip run make zip"
zip:
	powershell Compress-Archive -Force -Path .\dist\ -DestinationPath .\releases\Jira-Bridge-$(version).zip
installer: build-pyinstaller push-tags
	iscc /dMyAppVersion=$(version) .\setup-generator.iss
installer-debug: build-pyinstaller-debug
	@echo "run iscc with debug enabled"
	iscc /dMyAppVersion=$(version) .\setup-generator.iss

refresh-version:
	python deploy.py

push-tags:
	git push --tags

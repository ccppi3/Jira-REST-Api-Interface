version = 0.0.2
build-pyinstaller:
	echo "build using using pyinstaller"
	pyinstaller .\gui.py --add-data "jira.ico;." --add-data ".env;." --add-data ".\TemplatePayload.json;." --add-data ".\help.txt;." --hiddenimport win32timezone --icon .\jira.ico
	make zip


build-light:
	python -m nuitka --nofollow-import-to=pymupdf.mupdf --include-module=pymupdf.mupdf --show-scons --show-memory --show-modules --verbose --clang --jobs=9 --standalone --windows-console-mode=force --enable-plugin=tk-inter --show-progress --lto=no --include-data-files=.env=.env --include-data-files=jira.ico=jira.ico .\gui.py

zip:
	powershell Compress-Archive -Path .\dist\ -DestinationPath .\releases\Jira-Bridge-$(version).zip
build-good:
	python -m nuitka --show-scons --show-memory --show-modules --verbose --clang --jobs=9 --standalone --windows-console-mode=force --enable-plugin=tk-inter --show-progress --lto=no --include-data-files=.env=.env --include-data-files=jira.ico=jira.ico .\gui.py
	make zip



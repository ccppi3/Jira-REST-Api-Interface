version = 0.0.1

build-light:
	python -m nuitka --nofollow-import-to=pymupdf.mupdf --include-module=pymupdf.mupdf --show-scons --show-memory --show-modules --verbose --clang --jobs=9 --standalone --windows-console-mode=force --enable-plugin=tk-inter --show-progress --lto=no --include-data-files=.env=.env --include-data-files=jira.ico=jira.ico .\gui.py

zip:
	powershell Compress-Archive -Path .\gui.dist -DestinationPath .\releases\Jira-Bridge-$(version).zip
	python -m nuitka --standalone --windows-console-mode=force --enable-plugin=tk-inter --show-progress .\gui.py 
build-good:
	python -m nuitka --show-scons --show-memory --show-modules --verbose --clang --jobs=9 --standalone --windows-console-mode=force --enable-plugin=tk-inter --show-progress --lto=no --include-data-files=.env=.env --include-data-files=jira.ico=jira.ico .\gui.py
	make zip



version = 0.0.1

build-exe:
	python -m nuitka  --msvc=latest --standalone --windows-console-mode=force --enable-plugin=tk-inter --follow-imports --show-progress --nofollow-import-to=pymupdf.mupdf .\gui.py 
	make zip
zip:
	powershell Compress-Archive -Path .\gui.dist -DestinationPath .\releases\Jira-Bridge-$(version).zip
build-all:
	python -m nuitka --standalone --windows-console-mode=force --enable-plugin=tk-inter --show-progress .\gui.py 
build-debug:
	python -m nuitka --show-scons --show-memory --show-modules --verbose --clang --jobs=9 --standalone --windows-console-mode=force --enable-plugin=tk-inter --show-progress --lto=no .\gui.py 
	make zip



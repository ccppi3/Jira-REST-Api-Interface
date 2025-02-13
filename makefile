version = 0.0.1

build-exe:
		python -m nuitka --msvc=latest --standalone --windows-console-mode=disable --enable-plugin=tk-inter --follow-imports --show-progress .\gui.py 
	make zip
zip:
	powershell Compress-Archive -Path .\gui.dist - DestinationPath .\releases\Jira-Bridge-$(version).zip

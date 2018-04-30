
@echo on && echo Running python script && python hello_world.py && cd ..\Sphinx && echo Moving all rst files to Sphinx directory && copy /Y "..\Python Scripts\*.rst" ..\Sphinx && cd ..\Sphinx && echo Building Sphinx documentation from rsts && make html && cd "..\Python Scripts"

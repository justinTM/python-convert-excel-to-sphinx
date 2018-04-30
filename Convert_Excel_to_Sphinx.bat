
@echo on && echo Running python script && python populate_templates_from_sheets.py && cd ..\Sphinx && echo Moving all rst files to Sphinx directory && copy /Y "..\Python Scripts\*.rst" ..\Sphinx && cd ..\Sphinx && echo Building Sphinx documentation from rsts && make html && cd "..\Python Scripts"

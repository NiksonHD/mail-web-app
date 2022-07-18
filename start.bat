wkhtmltoimage_set_global_setting(settings, "loadPage.blockLocalFileAccess", "false");

del /q C:\Users\a\Music\NHD\python-scripts\input-files\*
copy "C:\Users\a\PRINT-WEB-SHOP\*" C:\Users\a\Music\NHD\python-scripts\input-files



C:\Users\a\AppData\Local\Programs\Python\Python310\python.exe "C:\Users\a\Music\NHD\python-scripts\make-order-last.py"
REM pause

rem Copies readme (and readme images) and TODO files to web dir
rem Copies release version of LdapQuery.dll to here


del web\index.html
del web\TODO.txt
copy readme.html web\index.html
@if ERRORLEVEL 1 GOTO kaboom
copy TODO.txt web
@if ERRORLEVEL 1 GOTO kaboom

del web\readme_images\*.*
md web\readme_images
copy readme_images\*.png web\readme_images
@if ERRORLEVEL 1 GOTO kaboom

del ldapquery.dll
copy ..\ldapquery\release\LdapQuery.dll .
@if ERRORLEVEL 1 GOTO kaboom

exit

:kaboom

echo "CRAP something went wrong"
pause


rem
rem Copies readme (and readme images) and TODO files to web dir
rem Copies release version of LdapQuery.dll to here
rem

del web\index.html
del web\TODO.txt
copy readme.html web\index.html
copy TODO.txt web

del web\readme_images\*.*
md web\readme_images
copy readme_images\*.png web\readme_images

del ldapquery.dll
copy ..\ldapquery\release\LdapQuery.dll .


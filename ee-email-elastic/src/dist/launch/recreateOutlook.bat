pushd %~dp0
call _settings.bat
%COMMAND% recreate outlook ee.email.outlook.ApplicationFactory
popd
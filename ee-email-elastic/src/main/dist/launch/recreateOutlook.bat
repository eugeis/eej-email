pushd %~dp0
call _settings.bat
%COMMAND% recreate outlook eugeis.email.outlook.ApplicationFactory
popd
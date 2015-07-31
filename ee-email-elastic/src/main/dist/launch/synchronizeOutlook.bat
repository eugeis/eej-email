pushd %~dp0
call _settings.bat
%COMMAND% synchronize outlook eugeis.email.outlook.ApplicationFactory
popd
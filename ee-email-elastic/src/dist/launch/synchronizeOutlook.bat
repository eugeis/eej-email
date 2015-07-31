pushd %~dp0
call _settings.bat
%COMMAND% synchronize outlook ee.email.outlook.ApplicationFactory
popd
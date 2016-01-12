echo off
rem Change DOS charset to UTF-8
chcp 65001

echo [INFO]Begin packaging...
call mvn clean dependency:copy-dependencies package -DoutputDirectory=lib -DincludeScope=compile -DskipTests=true  
echo [INFO]Package finished！
pause

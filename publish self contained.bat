dotnet publish -r win-x64 -p:PublishSingleFile=true --self-contained true --configuration release --output publish-sc /p:DebugType=None
pause
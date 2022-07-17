dotnet publish -r win-x64 -p:PublishSingleFile=true --self-contained false --configuration release --output publish /p:DebugType=None
pause
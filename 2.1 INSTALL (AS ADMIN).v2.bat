rem @echo off

@"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command "[System.Net.ServicePointManager]::SecurityProtocol = 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))" && SET "PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\bin"

choco install -y 7zip
rem choco install -y sharex
rem choco install -y copyq
rem choco install -y vscode
rem choco install -y git
rem choco install -y vlc
rem choco install -y python3
rem choco install -y wget
rem choco install -y windirstat
rem choco install -y cygwin
rem choco install -y gimp
choco install -y chocolateygui
choco install -y notepadplusplus
rem choco install -y processhacker
rem choco install -y libreoffice-fresh
choco install -y mattermost-desktop

rem call code --force --install-extension=alefragnani.project-manager
rem call code --force --install-extension=ms-vscode.atom-keybindings
rem call code --force --install-extension=ms-python.python
rem call code --force --install-extension=tht13.python
rem call code --force --install-extension=wmaurer.vscode-jumpy
rem call code --force --install-extension=alefragnani.bookmarks
rem call code --force --install-extension=helixquar.asciidecorator
rem call code --force --install-extension=sleistner.vscode-fileutils
rem call code --force --install-extension=editorconfig.editorconfig
rem call code --force --install-extension=kenhowardpdx.vscode-gist

rem echo [git] Configuring user aliases
rem git config --global alias.a "add -A"
rem git config --global alias.s "status"
rem git config --global alias.c "commit -m"
rem git config --global alias.p "push origin --all"
rem git config --global alias.adog "log --all --decorate --oneline --graph"
rem git config --global alias.fs "flow feature start"
rem git config --global alias.ff "flow feature finish"
rem git config --global alias.bs "flow bugfix start"
rem git config --global alias.bf "flow bugfix finish"
rem git config --global alias.hs "flow hotfix start"
rem git config --global alias.hf "flow hotfix finish"
rem git config --global alias.rs "flow release start"
rem git config --global alias.rf "flow release finish"

; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

[Setup]
AppName=Sebijk's Adressverwaltung
AppVerName=Sebijk's Adressverwaltung 1.07
AppCopyright=Copyright � 2005 - 2010 Home of the Sebijk.com. This software is released under the terms of the GNU General Public License.
AppPublisher=Home of the Sebijk.com
AppPublisherURL=http://www.sebijk.com
AppSupportURL=https://www.sebijk.com/community/index.php?board/13-home-of-the-sebijk-com-hilfe-unterstuetzung/
AppUpdatesURL=http://www.sebijk.com
VersionInfoCompany=Home of the Sebijk.com
VersionInfoDescription=Sebijk's Adressverwaltung
VersionInfoProductVersion=1.07
DefaultDirName={pf}\Home of the Sebijk.com\Adressverwaltung
DefaultGroupName=Home of the Sebijk.com\Adressverwaltung
AllowNoIcons=yes
InfoBeforeFile=X:\WorkingDir\Adressverwaltung\readme.txt
LicenseFile=X:\WorkingDir\Adressverwaltung\license.txt
OutputBaseFilename=adressen
SetupIconFile=X:\WorkingDir\Adressverwaltung\data\adressen.ico
Compression=lzma2
SolidCompression=true
WindowVisible=yes
DisableStartupPrompt=no
SignedUninstaller=yes

[Languages]
Name: "german"; MessagesFile: "compiler:Languages\German.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "X:\WorkingDir\Adressverwaltung\*"; DestDir: "{app}"; Flags: ignoreversion; Permissions: users-modify
Source: "X:\WorkingDir\Adressverwaltung\data\*"; DestDir: "{app}\data"; Flags: ignoreversion
Source: "X:\WorkingDir\Adressverwaltung\data\codebase\*"; DestDir: "{app}\data\codebase"; Flags: ignoreversion
Source: "X:\WorkingDir\Adressverwaltung\data\codebase\imgs\*"; DestDir: "{app}\data\codebase\imgs"; Flags: ignoreversion
Source: "X:\WorkingDir\Adressverwaltung\data\common\*"; DestDir: "{app}\data\common"; Flags: ignoreversion

[Icons]
Name: "{group}\Sebijk's Adressverwaltung"; Filename: "{app}\adressen.hta"; IconFileName: "{app}\data\adressen.ico"; WorkingDir: "{app}"
Name: "{group}\Hilfe zu Sebijk's Adressverwaltung"; Filename: "{app}\adressen.chm";
Name: "{group}\{cm:ProgramOnTheWeb,Home of the Sebijk.com}"; Filename: "http://www.sebijk.com"
Name: "{group}\Supportforum besuchen"; Filename: "https://www.sebijk.com/community/index.php?board/13-home-of-the-sebijk-com-hilfe-unterstuetzung/"
Name: "{group}\Adressdatenbank sichern"; Filename: "backup.vbs"; WorkingDir: "{app}";
Name: "{group}\{cm:UninstallProgram,Sebijk's Adressverwaltung}"; Filename: "{uninstallexe}";
Name: "{commondesktop}\Sebijk's Adressverwaltung"; Filename: "{app}\adressen.hta"; IconFileName: "{app}\data\adressen.ico"; WorkingDir: "{app}"; Tasks: desktopicon;
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\Sebijk's Adressverwaltung"; Filename: "{app}\adressen.hta"; IconFileName: "{app}\data\adressen.ico"; WorkingDir: "{app}"; Tasks: quicklaunchicon;

[Run]
Filename: "{app}\adressen.hta"; Description: "Sebijks Adressverwaltung starten"; Flags: nowait postinstall skipifsilent shellexec; WorkingDir: "{app}";
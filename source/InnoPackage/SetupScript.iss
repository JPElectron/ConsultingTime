[Setup]
AppName=Consulting Time
AppVerName=Consulting Time v2.9
PrivilegesRequired=admin
DefaultDirName={pf}\Consulting Time
DefaultGroupName=Consulting Time
AppCopyright=© 2010 JPElectron.com
AllowCancelDuringInstall=false
AllowUNCPath=false
ShowLanguageDialog=no
AppPublisher=JPElectron.com
AppPublisherURL=http://www.jpelectron.com
AppVersion=2.9
AppReadmeFile=http://www.jpelectron.com/readme/consultime.asp
UninstallDisplayIcon={app}\consult.exe
VersionInfoVersion=2.9
VersionInfoCompany=JPElectron.com
VersionInfoDescription=Consulting Time Setup
VersionInfoCopyright=(c) 2010 JPElectron.com
EnableDirDoesntExistWarning=false
DirExistsWarning=no
AppendDefaultGroupName=false
[Files]
Source: stdole2.tlb; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
Source: msvbvm60.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: oleaut32.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: olepro32.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: asycfilt.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile
Source: comcat.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: MCI32.OCX; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: consult.exe; DestDir: {app}; Flags: promptifolder
Source: Readme.url; DestDir: {app}
Source: Readme.url; DestDir: {group}
[Icons]
Name: {group}\Consulting Time; Filename: {app}\consult.exe; WorkingDir: {app}
[UninstallRun]
Filename: http://www.jpelectron.com/uf/frm.asp?type=new&app=consultime; Flags: shellexec
[Messages]
FinishedLabel=Setup has finished installing Consulting Time on your computer.
FinishedHeadingLabel=Consulting Time Setup Complete

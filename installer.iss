; Inno Setup Script for Voicemail Transcriber
; Build with: "C:\Users\PLM_Lacoursière\AppData\Local\Programs\Inno Setup 6\ISCC.exe" installer.iss

#define MyAppVersion "0.1.0"

[Setup]
AppName=Transcripteur Vocal
AppVersion={#MyAppVersion}
AppPublisher=Pierre-Luc Melancon
AppId={{B7E3D4A1-9F2C-4E8B-A5D6-1C3F7E9B2D4A}
DefaultDirName={autopf}\VoicemailTranscriber
DefaultGroupName=Transcripteur Vocal
OutputDir=installer_output
OutputBaseFilename=VoicemailTranscriber_Setup_{#MyAppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
SetupIconFile=app_icon.ico
UninstallDisplayIcon={app}\VoicemailTranscriber.exe
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
PrivilegesRequired=admin
DisableProgramGroupPage=yes

[Languages]
Name: "french"; MessagesFile: "compiler:Languages\French.isl"

[Files]
; Application files (PyInstaller output from voicemail_service.spec)
Source: "dist\VoicemailTranscriber\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; Supervisor exe (single-file PyInstaller output from voicemail_supervisor.spec)
Source: "dist\VoicemailSupervisor.exe"; DestDir: "{app}"; Flags: ignoreversion

; Whisper model files (bundled for offline use)
Source: "C:\Users\PLM_Lacoursière\.cache\huggingface\hub\models--mobiuslabsgmbh--faster-whisper-large-v3-turbo\snapshots\0a363e9161cbc7ed1431c9597a8ceaf0c4f78fcf\*"; DestDir: "{app}\models\faster-whisper-large-v3-turbo"; Flags: ignoreversion

[Icons]
Name: "{group}\Transcripteur Vocal"; Filename: "{app}\VoicemailSupervisor.exe"; IconFilename: "{app}\VoicemailTranscriber.exe"
Name: "{group}\Desinstaller Transcripteur Vocal"; Filename: "{uninstallexe}"

[Tasks]
Name: "startup"; Description: "Lancer automatiquement au demarrage de Windows"; GroupDescription: "Options supplementaires:"; Flags: checkedonce

[Registry]
; Auto-start at Windows login (if user checked the task). Target the
; supervisor so a native crash in the service is auto-recovered.
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "VoicemailTranscriber"; ValueData: """{app}\VoicemailSupervisor.exe"""; Flags: uninsdeletevalue; Tasks: startup

[Run]
Filename: "{app}\VoicemailSupervisor.exe"; Description: "Lancer Transcripteur Vocal maintenant"; Flags: nowait postinstall skipifsilent

[UninstallRun]
; Kill supervisor FIRST so it doesn't respawn the service mid-uninstall,
; then kill the service itself. Order matters.
Filename: "taskkill"; Parameters: "/F /IM VoicemailSupervisor.exe"; Flags: runhidden; RunOnceId: "KillSupervisor"
Filename: "taskkill"; Parameters: "/F /IM VoicemailTranscriber.exe"; Flags: runhidden; RunOnceId: "KillService"

[UninstallDelete]
; Clean up logs on uninstall
Type: filesandordirs; Name: "{app}\logs"

[Code]
procedure KillRunningProcesses;
var
  ResultCode: Integer;
begin
  // Supervisor first (so it doesn't respawn the service) then the service.
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM VoicemailSupervisor.exe',
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM VoicemailTranscriber.exe',
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  // Small delay so Windows releases file locks before we overwrite files
  Sleep(500);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssInstall then
    KillRunningProcesses;
end;

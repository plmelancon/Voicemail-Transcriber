# Dev-deploy shortcut (release path is the Inno Setup installer).
# Copies the service folder AND the supervisor exe into Program Files.
$src = "C:\CLAUDE CODE CLI PROJECTS\Voicemail Transcriber\dist\VoicemailTranscriber\*"
$srcSupervisor = "C:\CLAUDE CODE CLI PROJECTS\Voicemail Transcriber\dist\VoicemailSupervisor.exe"
$dst = "C:\Program Files\VoicemailTranscriber\"
Remove-Item -Path "$dst*" -Recurse -Force -ErrorAction SilentlyContinue
Copy-Item -Path $src -Destination $dst -Recurse -Force
Copy-Item -Path $srcSupervisor -Destination $dst -Force
Write-Host "Deploy complete."

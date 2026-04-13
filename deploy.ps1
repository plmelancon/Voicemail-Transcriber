$src = "C:\CLAUDE CODE CLI PROJECTS\Voicemail Transcriber\dist2\VoicemailTranscriber\*"
$dst = "C:\Program Files\VoicemailTranscriber\"
Remove-Item -Path "$dst*" -Recurse -Force -ErrorAction SilentlyContinue
Copy-Item -Path $src -Destination $dst -Recurse -Force
Write-Host "Deploy complete."

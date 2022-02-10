@echo off
dir /ad /b "TTS Output" 1>nul 2>nul && (echo Dir found.) || (md "TTS Output")
cls
main:
echo Welcome to the TTS System. All sounds are saved in .vbs files. Type what TTS you want it to be, and it'll be where your .bat is.
echo.
set /p tts=
echo Name of file must now be typed.
echo.
set /p name=
del create.vbs

echo Const SAFT48kHz16BitStereo = 39 >> create.vbs
echo Const SSFMCreateForWrite = 3 ' Creates file even if file exists and so destroys or overwrites the existing file >> create.vbs

echo Dim oFileStream, oVoice >> create.vbs

echo Set oFileStream = CreateObject("SAPI.SpFileStream") >> create.vbs
echo oFileStream.Format.Type = SAFT48kHz16BitStereo >> create.vbs
echo oFileStream.Open "./TTS Output/%name%.wav", SSFMCreateForWrite >> create.vbs

echo Set oVoice = CreateObject("SAPI.SpVoice") >> create.vbs
echo Set oVoice.AudioOutputStream = oFileStream >> create.vbs
echo oVoice.Speak "%tts%" >> create.vbs

echo oFileStream.Close >> create.vbs
echo dim filesys >> create.vbs
echo Set filesys = CreateObject("Scripting.FileSystemObject") >> create.vbs
echo If filesys.FileExists("./TTS Output/%name%.wav") Then filesys.DeleteFile "./create.vbs" >> create.vbs

start create.vbs

goto main

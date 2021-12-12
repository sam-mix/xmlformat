@ECHO OFF

pwd
go build -o text_replacer.exe .\cmd\text_replacer\main.go
text_replacer.exe
PAUSE

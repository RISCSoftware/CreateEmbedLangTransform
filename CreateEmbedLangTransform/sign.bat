@echo off
signtool sign /n "RISC Software GmbH" /a /fd SHA256 /tr http://ts.harica.gr/ /td SHA256 ..\Release\CreateEmbedLangTransform.exe
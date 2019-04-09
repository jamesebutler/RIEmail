rem Deploy application changes to Production server
robocopy .\bin\Debug\  \\S02AGPIAPP01\IP$\Reliability\BatchEmails\RIEmail.exe
robocopy .\bin\Debug\  \\S02AGPIAPP01\IP$\Reliability\BatchEmails\ *.dll
robocopy .\bin\Debug\  \\S02AGPIAPP01\IP$\Reliability\BatchEmails\ *.pdb
pause
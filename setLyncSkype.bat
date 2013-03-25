::
::  Simple BAT wrapper around powershell to get 32-bit powershell and make shortcut work nicely. Passes argument to powershell.
::
::  Drag to desktop as shortcut; right-click shortcut, select Properties; Add the word online after the text in Target:; Select the Shortcut Key: field, press CTRL+SHIFT+O; Click OK.
::  Again, drag to desktop as shortcut; right-click shortcut, select Properties; Add the word away after the text in Target:; Select the Shortcut Key: field, press CTRL+SHIFT+A; Click OK.
::
::  Orig Author/Date/(c): awcoleman/20130325/awcoleman    License: MIT (http://opensource.org/licenses/mit-license.html)
::  Last Update: awcoleman/20130325
::
C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe C:\Users\acoleman\setLyncSkype.ps1 %1
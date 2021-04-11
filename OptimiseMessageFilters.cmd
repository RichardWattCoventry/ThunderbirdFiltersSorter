@echo off
@echo Making a backup copy of the Thunderbird message filters file...
@rem xcopy "C:\Users\Richard\Documents\Emails\Thunderbird\msgFilterRules.dat" "C:\Users\Richard\Documents\Emails\msgFilterRules.bak" /Y/R/F/I
@rem xcopy Source Target
@rem Source: C:\Users\Richard\Online\Emails\Thunderbird\msgFilterRules.dat
@rem Target: C:\Users\Richard\Online\Emails\msgFilterRules.bak
xcopy "C:\Users\Richard\Online\Emails\Thunderbird\msgFilterRules.dat" "C:\Users\Richard\Online\Emails\msgFilterRules.bak" /Y/R/F/I
@echo Thunderbird message filters file backed up.
@echo Sorting the Thunderbird message filters file...
@rem c:\Windows\SysWOW64\cscript.exe //nologo "C:\Users\Richard\Documents\Developer\Scripts\VBScript\Message Filters\MessageFilters.vbs" "C:\Users\Richard\Documents\Emails\Thunderbird\msgFilterRules.dat"
@rem C:\Windows\SysWOW64\cscript.exe //nologo Script Source
@rem Script: "C:\Users\Richard\Programming\Scripts\VBScript\Message Filters\MessageFilters.vbs"
@rem Source: "C:\Users\Richard\Online\Emails\Thunderbird\msgFilterRules.dat"
c:\Windows\SysWOW64\cscript.exe //nologo "C:\Users\Richard\Programming\Scripts\VBScript\Message Filters\MessageFilters.vbs" "C:\Users\Richard\Online\Emails\Thunderbird\msgFilterRules.dat"
@echo Thunderbird message filters file sorted.
:end

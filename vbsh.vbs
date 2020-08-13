' @author: Grok
' @license: MIT
'
' A simple interactive VBScript shell.

version = "0.11"

' Add all scripts you want to import at startup
scripts = array()

' Set to True if you want to auto import files in scripts array
allways_import_scripts = False

' Default logging level is INFO
LogLevel = 4
LogPrefix = True


ParseInputArgs()
Boot()


private function Boot()
    LogDebug("Boot()")

    if allways_import_scripts = True then
        ImportInitScript()
        Main()
    else
	if ubound(scripts) >= 0 then
		' Only prompt user input if there is any predefined scripts
		LogInfo("Do you want to import the predefined scripts? (In this order)" & VbCrLf & "")
		for each script in scripts
		LogInfo(" * " + script)
		next

		wscript.echo ""
		wscript.stdout.write("(y / n / q): ")

		line = trim(wscript.stdin.readline)

		select case line
			case "y"
				ImportInitScript()
				Main()
			case "n"
				Main()
			case "q" 
				Print("Bye :]")
				wscript.quit(0)
			case else
				LogError("Invalid input... Try again!")
		end select
        else
            Main()
        end if
    end if
end function 'private function Boot()



private function ImportInitScript()
    LogDebug("ImportInitScript()")

    for each script in scripts
        LogDebug("Import Initscript: " + script)

        set sh  = createobject("WScript.Shell")
        set fso = createobject("Scripting.FileSystemObject")

        path = sh.ExpandEnvironmentStrings(script)
        scriptExists = fso.FileExists(path)

        Set sh  = Nothing
        Set fso = Nothing

        if scriptExists then
            Import(path)
        end if
    next
end function


private function Main()
    LogDebug("Main()")

    do while True
        wscript.stdout.write(">>> ")

        line = trim(wscript.stdin.readline)
        do while right(line, 2) = " _" or line = "_"
            line = rtrim(left(line, len(line)-1)) & " " & trim(wscript.stdin.readline)
        loop

        select case StartsWith(line)
		case "?exit" , "?q" , "?quit"
			exit do
		case "?"
			PrintHelp()
		case "?import "
			file = Replace(line, "?import ", "")
			Import(file)
		case "?version"
			Print(version)
		case "?reimport"
			ImportInitScript()
		case "?config"
			PrintHelpConfig()
		case "?run"
			PrintHelpRun()	
		case "?ext"
			PrintHelpExtractions()
		case "?ini"
			showIniValues()
		case "sei"
			sei(line)
		case else
			on error resume next
			Err.clear
			Execute(line)
			if Err.Number <> 0 then
			fedcba = trim(Err.Description & " (0x" & hex(Err.Number) & ")")
				if Err.Number = 13 then
					abcdef = "if VarType(" & line & ") = 0 then" & VbCrLf & _
						 "     wscript.echo " & quotes("Object not initilized") & VbCrLf & _
						 "else" & VbCrLf & _
						 "     wscript.echo CStr(" & line & ")" & VbCrLf & _
						 "end if" & VbCrLf
					ExecuteCode(abcdef)
				else
					wscript.echo "Compile-Error: " + Err.Description
				end if
			end if
			on error goto 0
	end select
    loop
end function 'private function Main()


private function ExecuteCode(ByRef s)
    LogDebug("ExecuteCode()")
    LogDebug(s)
    Execute(s)
end function


private function quotes(string)
    LogDebug("quotes(" & string & ")")
    quotes = chr(34) + string + chr(34)
end function


private function PrintHelp()
    LogDebug("PrintHelp()")

    ' Print a Help message.
    Print("VBSchell " + version)
    Print("")
    Print("   ?           Prints this Help")
    Print("   ?exit       Exit the shell")
    Print("   ?import     Prompts to import a .vbs script")
    Print("   ?reimport   Reimports the predefined scripts")
    Print("   ?version    Prints the version")
end function


' Import the first occurrence of the given filename from the working directory
' or any directory in the %PATH%.
'
' @param  filename   Name of the file to import.
private function Import(ByVal filename)
    LogDebug("Import(" & filename & ")")

    set fso = createobject("Scripting.FileSystemObject")
    set sh = createobject("WScript.Shell")

    filename = trim(sh.ExpandEnvironmentStrings(filename))
    If Not (left(filename, 2) = "\\" or mid(filename, 2, 2) = ":\") then
        ' filename is not absolute
        if not fso.FileExists(fso.GetAbsolutePathName(filename)) then
            ' file doesn't exist in the working directory => iterate over the
            ' directories in the %PATH% and take the first occurrence
            ' if no occurrence is found => use filename as-is, which will result
            ' in an error when trying to open the file
            for each dir in split(sh.ExpandEnvironmentStrings("%PATH%"), ";")
                if fso.FileExists(fso.BuildPath(dir, filename)) then
                    filename = fso.BuildPath(dir, filename)
                    exit for
                end if
            next
        end if
        filename = fso.GetAbsolutePathName(filename)
    end if

    if fso.FileExists(filename) then
        set file = fso.OpenTextFile(filename, 1, False)
        code = file.ReadAll()
        file.Close()

        LogInfo("Importing file: " + filename)
        ExecuteGlobal(code)
    else
        LogError("File Not Found on disk...")
    end if

    set fso = Nothing
    set sh = Nothing
end function


private function StartsWith(string, what)
    LogDebug("StartsWith(" & string & ", " & what & ")")

    if InStr(Trim(string), what) = 1 then
        StartsWith = True
    else
        StartsWith = False
    end if

    LogDebug("return StartsWith() = " & StartsWith)
end function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Input argument parsing
private function ParseInputArgs()
LogDebug("ParseInputArgs()")

    ' Parse all input arguments and set variables
    for i = 0 to wscript.Arguments.count - 1
        select case wscript.arguments.Item(i)
			case "--help"
				' Using docopt cli specs "https://github.com/docopt/docopt"
				Print("Usage:")
				Print(" vbsh.wsf [-v ... ] [--help]")
				Print(" vbsh.wsf --version")
				Print(" ")
				Print("Options:")
				Print(" --help      Prints this help")
				Print(" -v ...      Sets the logging level of this run. [Default: -vvvv --> INFO logging] [Supports level 1-5]")
				Print(" --version   Prints the version of the application")
				Print("")
				wscript.quit(0)
			case "--version" 
				wscript.echo version
				wscript.quit(0)
			case "-v"
				' Critical log level
				LogLevel = 1
			case "-vv"
				' Error log level
				LogLevel = 2
			case "-vvv"
				' Warning log level
				LogLevel = 3
			case "-vvvv"
				' Info log level
				LogLevel = 4
			case "-vvvvv"
				' Debug log level
				LogLevel = 5
			case else
				wscript.echo "ERROR: Unknown input argument: " + a
				wscript.quit(0)
        end select
    next
end function 'private function ParseInputArgs()


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Logging functions'

private function LogDebug(msg)
    LogInternal 5, "DEBUG: ", msg
end function

private function LogInfo(msg)
    LogInternal 4, "INFO: ", msg
end function

private function LogWarning(msg)
    LogInternal 3, "WARNING: ", msg
end function

private function LogError(msg)
    	LogInternal 2, "ERROR: ", msg
end function

private function LogCritical(msg)
	LogInternal 1, "CRITICAL: ", msg
end function

private function LogInternal(preLogLevel, preString, msg)
    if LogLevel >= preLogLevel then
	if LogPrefix = True then
            print(preString & cstr(msg))
        else
            print(cstr(msg))
        end if
    end if
end function

private function print(msg)
    wscript.echo msg
end function

private function p(msg)
    wscript.echo msg
end function

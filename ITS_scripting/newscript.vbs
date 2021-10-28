'New asp - read a template and fill in some blanks by asking some questions
option explicit 
Dim fso
dim oFile
dim i



const template="templatev7.wsf"


main

sub main

	dim tso1
	dim tso2
	dim userArray
	dim newscriptname
	dim oneliner
	dim YourAPieceOfString


	set FSO = wscript.createobject("scripting.filesystemobject")
	
	if fso.fileexists(template) then
		
		set tso1 = fso.opentextfile(template)
		
		YourAPieceOfString = tso1.readall
		
		tso1.close
		
		set tso1 = nothing
		
	else
	
		wscript.echo "Can't find " & template
		
	end if
	
	 
	newscriptname = inputbox ("New script name (no extenstions)-->", "Create a new script")
	
	if newscriptname = "" then exit sub
	
	newscriptname = newscriptname & ".wsf"
	
	
	if fso.fileexists(newscriptname) then
		wscript.echo "Sorry, that file exists " & newscriptname
		exit sub
	end if
	
	oneliner = inputbox ("Enter a very brief description-->", "One Line description")
	
	YourAPieceOfString = replace(YourAPieceOfString,"script_name_token", newscriptname)
	YourAPieceOfString = replace(YourAPieceOfString,"date_token", now())
	YourAPieceOfString = replace(YourAPieceOfString,"one_line_token", oneliner)
	
	wscript.echo "Creating " & newscriptname
	
	set tso2 = fso.createtextfile(newscriptname)
	
	tso2.write YourAPieceOfstring
	
	tso2.close
	

end sub


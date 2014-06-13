<%
	Const cmd_help = "WELCOME TO MY HOMEPAGE!<br>You can open post by 'open [title]' command.<br>In addition, the screen can be cleaned with 'clear' command."
	Const cmd_clear = "The screen have been cleaned."
	Const cmd_open = "You should write 'open [title]'."
	Const cmd_noone = "It is not a command."
	cmd = request.Form("command")
	argument = request.Form("mode")
	tempstr = ""
	If (IsEmpty(Session("log"))) then Session("log") = ".\><font>help</font><br><div class=" & Chr(34) & "postbox" & Chr(34) & ">" & cmd_help & "</div><br>"
	If (argument="write") then
		If (Right(LCase(cmd), 8)=" & clear") then
			cmd = Left(cmd, Len(cmd)-8)
			tempstr = " & clear"
			Session("log") = ""
		End If
		If (LCase(Rtrim(cmd))="help") then
			post = cmd_help
			filename = "noone"
		ElseIf (LCase(Rtrim(cmd))="clear") then
			post = cmd_clear
			filename = "noone"
			Session("log") = ""
		ElseIf (LCase(Rtrim(cmd))="open") then
			post = cmd_open
			filename = "noone"
		ElseIf (LCase(cmd)="open 3dp 칩") then
			filename = "bottom_3dpchip.html"
		ElseIf (LCase(cmd)="open 반디집") then
			filename = "bottom_bandizip.html"
		ElseIf (LCase(cmd)="open 꿀뷰") then
			filename = "bottom_honeyview.html"
		ElseIf (LCase(cmd)="open 알송") then
			filename = "bottom_alsong.html"
		ElseIf (LCase(cmd)="open 팟플레이어") then
			filename = "bottom_daumpotplayer.html"
		Else
			post = "'" & cmd & "' " & cmd_noone
			filename = "noone"
		End If
		If filename<>"noone" then
			Set stream = Server.CreateObject("ADODB.Stream")
			With Stream
				.Type = 2
				.CharSet = "UTF-8"
				.Open
				.LoadFromFile(Server.MapPath(".\" & filename))
				post = .ReadText()
				.Close
			End With
		End If
		Session("log") = Session("log") & ".\><font>" & cmd & tempstr & "</font><br><div class=" & Chr(34) & "postbox" & Chr(34) & ">" & post & "</div><br>"
	End If
%>
<!doctype html>
<html>
    <head>
        <meta charset="utf-8">
        <title>Bottom</title>
        <style type="text/css">
			html { height:100%; }
			body { margin:0px; height:100%; background-color:black; color:#E4E4E4; font-family:"Courier New"; }
			.box { border-style:solid; border-width:0px 2px 0px 0px; border-color:rgb(223,237,252); height:100%; padding-left:16px; }
			.cmdtxt {
				background-color:black;
				border-style:solid;
				border-width:1px;
				border-color:rgb(38,72,127);
				border-radius:3px;
				color:rgb(143,239,143);
				font-family:"Courier New";
			}
			.subtxt {
				background-color:black;
				border-width:0px;
			}
			font { color:rgb(143,239,143); }
			.postbox { border-style:solid; border-width:1px; border-color:#FF6E3E; border-radius:3px; padding:5px; width:90%; }
			a:link    { color:#FF8E00; text-decoration:none; }
			a:visited { color:#FF8E00; text-decoration:none; }
			a:hover   { color:#FF8E00; text-decoration:underline; }
			a:active  { color:#FF8E00; text-decoration:none; }
		</style>
        <link rel="stylesheet" type="text/css" href="scrollbar.css">        
        <script type="text/javascript" src="scrollbar_jquery.min.js"></script>
        <script type="text/javascript" src="scrollbar.js"></script>
        <script type="text/javascript" src="scrollbar_main.js"></script>
        <script type="text/javascript" src="rainbow.js"></script>
    </head>
    <body>
    	<div class="box">
        <div class="nano">
        <div class="nano-content">
        	<br>JunguSoft Windows [Version 1.0.0]<br>
            Copyright (c) 2014 JunguSoft Corporation. All rights reserved.<br><br>
            <% =Session("log") %>
            <form name="txtcmd" action="bottom_main.asp" method="post">
            	<% If (argument="input") then %>
            		.\><input type="text" name="command" class="cmdtxt" value="<% =cmd %>" size="60px" autocomplete="off" autofocus required>
                <% Else %>
                	.\><input type="text" name="command" class="cmdtxt" size="60px" autocomplete="off" autofocus required>
                <% End If %>
                <input type="hidden" name="mode" value="write">
                <input type="submit" class="subtxt" value="">
            </form>
            <br>
        </div>
        </div>
        </div>
    </body>
</html>

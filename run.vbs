Set Shell = CreateObject("WScript.Shell")

     Answer = MsgBox("Do you wanna update dependecies?",vbYesNo,"Dependecies")
     If Answer = vbYes Then
          Shell.run "npm i"
            Shell.run "node index.js"
          Ending = 1
     ElseIf Answer = vbNo Then
         Answer = MsgBox("Ok.Wanna start the bot?",1+16,"Bot Starting")
         Shell.run "node index.js"
         End If
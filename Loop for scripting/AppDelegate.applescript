--
--  AppDelegate.applescript
--  Loop for scripting
--
--  Created by Jonathan Stoff on 8/3/21.
--  
--
script AppDelegate
    use framework "Cocoa"
    use framework "Appkit"
    use framework "foundation"
    use framework "ScriptingBridge"
    use scripting additions
    use framework "ApplicationServices"
    property parent : class "NSObject"
    property NSPopUpButton : class "NSpopUpButton"
    #property theadderar : {}
    property NSMenuItem : class "NSMenuItem"
    property NSButton1 : class "NSButton"
    property NSImage : class "NSImage"
    property NSUserDefaults : class "NSUserDefaults"
    property NSColor : class "NSColor"
    property Appliz : current application
    property color0 : missing value
    property standardUserDefaults : missing value
    property statusMenu : missing value
    property dynamicMenu : missing value
    property textbox1 : missing value
    property textbox2 : missing value
    property searchstr : missing value
    property textbox3 : missing value
    property buttoncol1 : missing value
    property textsearchs1 : missing value
    property progressbar1 : missing value
    property statusItemController : missing value
	property theWindow : missing value
    property switch1 : missing value
    property listobj1 : missing value
    property theconz : true
    property datalinefp : ""
    property r60cb1 : missing value
    property r30cb1 : missing value
    property t30cb1 : missing value
    property t15cb1 : missing value
    property t10cb1 : missing value
    property t60cb1 : missing value
    property rs60cb1 : missing value
    property rs30cb1 : missing value
    property ts30cb1 : missing value
    property ts15cb1 : missing value
    property ts10cb1 : missing value
    property ts60cb1 : missing value
    property r60cb2 : missing value
    property r30cb2 : missing value
    property t30cb2 : missing value
    property t15cb2 : missing value
    property t10cb2 : missing value
    property t60cb2 : missing value
    property ctwin1 : missing value
    property ctwin2 : missing value
    property ctwin3 : missing value
    property printer : true
    property switchp : missing value
    property ct3button : missing value
    property dynamicMenuwin : missing value
    property dynamicTypeMenu : missing value
    property lastfivespp : {}
    property lastfiveijp: {}
    property lastfivesjp : {}
    property lastfivetvjp : {}
    property lastsessnamep : ""
    property sessnameszlap : {}
    property firsttimvp : true
    property psessdir : posix file ("/Users/jonathanstoff/Documents/macrostuff/Previous_sessions.txt")
    property zsessdir : posix file ("/Users/jonathanstoff/Documents/macrostuff")
    on printsetting_(sender)
        if ((my switchp)'s state) as string = "1" then
            set (my printer) to true
        else if ((my switchp)'s state) as string = "0" then
            set (my printer) to false
        end if
    log (my switchp)'s state as string
    log (my printer) as string
    end printsetting_
    on applicationWillFinishLaunching_(aNotification)
        try
      #my ctwin1's orderOut_(sender)
      #my ctwin2's orderOut_(sender)
      log "loading"
        set menuitemz to (my dynamicMenu)
        set typemenu to (my dynamicTypeMenu)
        menuitemz's removeAllItems()
        typemenu's removeAllItems()
        tell application "Finder"
            set tempdirz to my psessdir
            set tempfil to my readFile("/Users/jonathanstoff/Documents/macrostuff/Previous_sessions.txt")
            if tempfil is "" then
            else
            set tempfil to my trimThis(tempfil, true, "full")
            set temparf to my theSplit(tempfil, "*")
            set the_count to (count temparf) - 1
            set in_t1 to 1
            repeat the_count times
                set str_z to (item in_t1 of temparf)
                menuitemz's addItemWithTitle_(str_z)
                set in_t1 to in_t1 + 1
            end repeat
            end if
            end tell
        set the_action to ((my dynamicMenu's titleOfSelectedItem) as Unicode text)
        set datalinee to my menutocellnum(the_action)
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                    tell sheet "Main Base"
                        set thezell to cell ("G" & datalinee)
                        set thetypse to my theSplit(value of thezell, "*")
                    end tell
            end tell
        end tell
            typemenu's addItemsWithTitles_(thetypse)
        log tempfil
        tell application "System Events"
            set UI_enabled to UI elements enabled
        end tell
        if UI_enabled is false then
        tell application "System Preferences"
            activate
            set securityPane to pane id "com.apple.preference.security"
            tell securityPane to reveal anchor "Privacy_Accessibility"
            #set current pane to pane id "com.apple.preference.universalaccess"
        end tell
        end if
        end try
    end applicationWillFinishLaunching_
	on applicationShouldTerminate_(sender)
		return current application's NSTerminateNow
	end applicationShouldTerminate_
    on changetypes_(sender)
        set typemenu to (my dynamicTypeMenu)
        typemenu's removeAllItems()
        set the_action to ((my dynamicMenu's titleOfSelectedItem) as Unicode text)
        set datalinee to my menutocellnum(the_action)
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                    tell sheet "Main Base"
                        set thezell to cell ("G" & datalinee)
                        set thetypse to my theSplit(value of thezell, "*")
                    end tell
            end tell
        end tell
        typemenu's addItemsWithTitles_(thetypse)
    end changetypes_
    on pcshutterdown_(sender)
        tell application "All-in-One Messenger" to quit
        tell application "Microsoft Word" to quit
        tell application "Microsoft Outlook" to quit
        tell application "App Store" to quit
        tell application "Find Any File" to quit
        tell application "Notes" to quit
        tell application "Microsoft Teams" to quit
        tell application "System Events"
            try
            tell process "Pro Tools"
                set frontmost to true
                delay 1
                click menu item "Save" of menu "File" of menu bar 1
            end tell
            end try
        end tell
        tell application "TextMate" to quit
            tell application "Microsoft Excel"
                tell workbook "Database.xlsx"
                    Save
                end tell
                tell workbook "Audio Production Sheet.xlsx"
                    Save
                end tell
            end tell
            tell application "Keyboard Maestro Engine"
                do script "Doublecheckshutdown"
            end tell
            quit
    end pcshutterdown_
    on REMOVENAMES_(sender)
    tell application "Microsoft Excel"
        tell workbook "Audio Production Sheet.xlsx"
            tell sheet "Audio"
            set tempnumj to 3
            repeat 50 times
                set cellsell to "A" & tempnumj
                set cellzone to value of cell cellsell
                log "got value"
                set cellvalone to cellzone
                if cellzone contains "(Jerry)" then
                    set cellvalone to (my theSplit(cellzone, "(Jerry)") as string)
                else if cellzone contains "(Max)" then
                    set cellvalone to (my theSplit(cellzone, "(Max)") as string)
                end if
                set value of cell ("A" & tempnumj) to cellvalone
                set tempnumj to tempnumj + 1
            end repeat
            end tell
        end tell
    end tell
    end REMOVENAMES_
    on startupt_(sender)
    tell application "Keyboard Maestro Engine"
        do script "Starup"
    end tell
    end startupt_
    on ctwin1but_(sender)
        set my theconz to false
    end ctwin1but_
    on testswitch_(sender)
        performSelectorInBackground_withObject_("removesess", "")
    end testswitch_
    on sendoutprocsh_(sender)
        performSelectorInBackground_withObject_("prosentout", "")
    end sendoutprocsh_
    on prosentout()
        ignoring application responses
        tell application "Keyboard Maestro Engine"
            do script "Send out production script"
        end tell
        end ignoring
    end prosentout
    on addreplyemail_(sender)
        performSelectorInBackground_withObject_("addreplzemail", "")
    end addreplyemail_
    on addreplzemail()
        tell application "Microsoft Outlook"
            set themesss to selection
            set themesss to item 1 of (get current messages)
            set subbyx to subject of themesss as string
            my theReplyz(themesss, subbyx)
        end tell
        
    end addreplzemail
    on openurlMusurl_(sender)
        tell application "Google Chrome" to open location "https://www.universalproductionmusic.com/en-us/login"
            tell application "Keyboard Maestro Engine"
                do script "7C8495FE-EDE5-43E6-ABAE-9F73AD59A3EB"
            end tell
            delay 2
            tell application "Google Chrome"
                set thetab to active tab of front window
                set url of thetab to "https://www.universalproductionmusic.com/en-us"
            end tell
    end openurlMusurl_
    on openurlSFXurl_(sender)
         tell application "Google Chrome" to open location "https://www.universalproductionmusic.com/en-us/login"
        tell application "Keyboard Maestro Engine"
            do script "7C8495FE-EDE5-43E6-ABAE-9F73AD59A3EB"
        end tell
        delay 2
        tell application "Google Chrome"
            set thetab to active tab of front window
            set url of thetab to "https://upmsfx.sounddogs.com/en-us/?key=jNFFKbbs1CK7ih0e25FrNT0qMgsq5JfEyLoM4vmkPSF%2BAhwWbDSE%2BR8Z9K9EBUov1pz8bl%2BVvH3ECzfbo7qUn2c3%2FfIo6AsMKjaT%2BOgi5d4kshRuwM39n2JB2%2Fnqh%2BEyLr031U3MyqCv%2BSlFf5nyLw%3D%3D&source=cw2"
        end tell
    end openurlSFXurl_
    on searchSFXfold_(sender)
        performSelectorInBackground_withObject_("searchSFXfold", " ")
    end searchSFXfold_
    on openoldsessfo_(sender)
        set the_action to my textsearchs1's stringValue as string
        set datalinee to my menutocellnum(the_action)
        log "openoldsessfo var: " & datalinee
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
                    set thecellr to (cell ("C" & datalinee))
                    set thebesh to value of thecellr as string
                    
                end tell
            end tell
        end tell
        tell application "Finder" to open (thebesh as alias)
    end openoldsessfo_
    on searchSFXfold()
        set searchtermt1 to my searchstr's stringValue as string
        set searchfafurl to "fafapp://find?inp1=" & searchtermt1 & "&loc=/Volumes/JDA_DATA/MAIN SFX"
        open location searchfafurl
    end searchSFXfold
    on searchMUSfold_(sender)
        performSelectorInBackground_withObject_("searchMUSfold", " ")
    end searchMUSfold_
    on searchMUSfold()
        set searchtermt1 to my searchstr's stringValue as string
        set searchfafurl to "fafapp://find?inp1=" & searchtermt1 & "&loc=/Volumes/JDA_DATA/MUSIC LIBRARY"
        open location searchfafurl
    end searchMUSfold
    on searchthedata_(sender)
        set beh to true
        #my progressbar1's startAnimation_(1)
        set textsearchss1 to (my textsearchs1)
        set searchterm to textsearchss1's stringValue as string
        if searchterm is "" then
        else if searchterm is " " then
        else if searchterm is "  " then
        else
        if searchterm contains "<" then
           set datalinee to my menutocellnum(searchterm)
           set my textbox3's stringValue to datalinee as string
        end if
        if searchterm contains "|" then
            set thereals to my theSplit(searchterm, " | ")
            set searchterm to (item 1 of thereals) & " " & (item 2 of thereals)
        end if
        my textsearchs1's removeAllItems()
        my textsearchs1's reloadData()
            set theadderar to my searchydofu(searchterm)
            
        my textsearchs1's addItemsWithObjectValues_(theadderar)
        my textsearchs1's reloadData()
        end if
        #my progressbar1's stopAnimation_(1)
    end searchthedata_
    on searchydofu(searchterm)
        tell application "Microsoft Excel"
                   tell workbook "Database.xlsx"
                       tell sheet "Main Base"
                           set theadderar to {}
                           set bzz to true
                           set rangz to {}
                           set tcnumin to 2
                           
                           repeat while bzz is true
                                  try
                               set rangez to find range ("A" & tcnumin & ":H" & 3100) what searchterm
                               set tcnum to (first row index of rangez)
                               if rangz contains tcnum then
                                   exit repeat
                               else
                               set end of rangz to tcnum
                               end if
                               set tcnumin to tcnum + 1
                           on error
                               exit repeat
                           end try
                           end repeat
                           
                       repeat with cnum in rangz
                           set r1f to cell ("F" & cnum)
                           set r2E to cell ("E" & cnum)
                           set r3D to cell ("D" & cnum)
                           set r4A to cell ("A" & cnum)
                           if value of r1f is "" then
                               set theadder to (value of r4A) & " | <" & cnum & ">"
                            else
                           set theadder to (value of r1f) & " | " & (value of r3D) & " | " & (value of r2E) & " | <" & cnum & ">"
                           
                           end if
                           set end of theadderar to theadder
                           end repeat
                       end tell
                   end tell
               end tell
        return(theadderar)
    end searchydofu
    #on searchthedata_(sender)
        #performSelectorInBackground_withObject_("searchbackdat", "")
        #my searchbackdat()
    #end searchthedata_
    on adddatalinetocurrent_(sender)
        set tempar1 to {}
        set textybox2 to (my textbox3)
        set cellnumber to textybox2's stringValue as string
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
                    set tempran1 to "F" & cellnumber
                    if (value of cell tempran1) as string is "" then
                        set tempran5 to "A" & cellnumber
                        set tempstr1 to value of cell tempran5 & " | " & "<" & cellnumber & ">"
                    else
                    set tempran2 to "D" & cellnumber
                    set tempran3 to "E" & cellnumber
                    set tempstr1 to value of cell tempran1 & " | " & value of cell tempran2 & " | " & value of cell tempran3 & " | " & "<" & cellnumber & ">"
                    end if
                end tell
            end tell
        end tell
        my writetopsess(tempstr1)
    end adddatalinetocurrent_
    on dupsess_(sender)
        performSelectorInBackground_withObject_("dupsessr", "")
    end dupsess_
    on dupsessr()
        set textybox3 to (my textbox3)
        set datalinee to textybox3's stringValue as string
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
                    set ransess to "C" & datalinee
                    set foldersees to value of cell ransess
                    
                end tell
            end tell
        end tell
        set partem to my trimThis(foldersees, true, "full")
        set parar to my theSplit(foldersees, ":")
        if item -1 of parar is " " then
        set counm to (count parar) - 2
        else if item -1 of parar is "" then
        set counm to (count parar) - 2
        else
        set counm to (count parar) - 1
        end if
        set parzf to {}
        set int22 to 1
        repeat counm times
            set end of parzf to (item int22 of parar) & ":"
            set int22 to int22 + 1
        end repeat
        set parentfol to parzf as string
        set the_action to ((my dynamicMenu's titleOfSelectedItem) as Unicode text)
        set cellnumber to my menutocellnum(the_action)
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
                    set thrdh to value of (cell ("L" & cellnumber))
                    set allauf to my theSplit(thrdh, "*")
                    set ransess to "H" & cellnumber
                    set docxH to value of cell ransess
                    try
                        set docxHar to my theSplit(docxH, "*")
                        repeat with doc1 in docxHar
                            if doc1 contains "60" then
                                set docxH to doc1
                                exit repeat
                            else if doc1 contains "TV"
                                if doc1 contains "30"
                                    set docxH to doc1
                                    exit repeat
                                end if
                            end if
                        end repeat
                    end try
                    set ransess to "F" & cellnumber
                    set ranzsess to "D" & cellnumber
                    set filnamez to (value of cell ransess) & " " & (value of cell ranzsess)
                    set theranz to "A" & cellnumber
                    set value of cell theranz to filnamez & ".ptx"
                    set theranz to "B" & cellnumber
                    set value of cell theranz to parentfol & filnamez & ":" & filnamez & ".ptx"
                    set theranz to "C" & cellnumber
                    set value of cell theranz to parentfol & filnamez & ":"
                    my setsessmoddate(cellnumber)
                end tell
            end tell
        end tell
        tell application "Finder"
            try
            set FolderCopy to duplicate foldersees
            end try
            set sortedff to get every item of FolderCopy
            set countyo to count of sortedff
            set inttyy to 1
            repeat countyo times
                set checkzf to (item inttyy of sortedff) as string
                try
                if (name of checkzf) contains "mp3" then
                    delete (item inttyy of sortedff)
                else if (name of checkzf) contains "wav" then
                    delete (item inttyy of sortedff)
                else if (name of checkzf) contains "aif" then
                    delete (item inttyy of sortedff)
                else
                set commandydoo to do shell script "/usr/local/Cellar/tag/0.10_1/bin/tag -r red " & "\"" & POSIX path of (item inttyy of sortedff as text) & "\""
                set thenewcommand to do shell script "/Library/Developer/CommandLineTools/usr/bin/SetFile -a l " & "\"" & POSIX path of (item inttyy of sortedff as text) & "\""
                if checkzf contains ".doc" then
                    delete (item inttyy of sortedff)
                else if checkzf contains "ptx" then
                    set name of item inttyy of sortedff to filnamez & ".ptx"
                end if
                end if
                on error
                set commandydoo to do shell script "/usr/local/Cellar/tag/0.10_1/bin/tag -r red " & "\"" & POSIX path of (item inttyy of sortedff as text) & "\""
                set thenewcommand to do shell script "/Library/Developer/CommandLineTools/usr/bin/SetFile -a l " & "\"" & POSIX path of (item inttyy of sortedff as text) & "\""
                if checkzf contains ".doc" then
                    delete (item inttyy of sortedff)
                else if checkzf contains "ptx" then
                    set name of item inttyy of sortedff to filnamez & ".ptx"
                end if
                end try
                set inttyy to inttyy + 1
            end repeat
            if docxH contains "/" then
                set theDocx to docxH as posix file
                move file theDocx to FolderCopy
            else if docxH contains ":"
                set theDocx to docxH as alias
                move file theDocx to FolderCopy
            end if
            try
            repeat with audiofiz in allauf
                if audiofiz contains ":" then
                    set thef to audiofiz as alias
                    set commandydoo to do shell script "/usr/local/Cellar/tag/0.10_1/bin/tag -s red " & "'" & posix path of (thef) & "'"
                    move file thef to FolderCopy
                else if audiofiz contains "/" then
                    set thef to audiofiz as posix file
                    set commandydoo to do shell script "/usr/local/Cellar/tag/0.10_1/bin/tag -s red " & "'" & thef & "'"
                    move file thef to FolderCopy
                end if
            end repeat
            end try
            set name of FolderCopy to filnamez
        end tell
    end dupsessr
    on Movevotosess_(sender)
        set menuitemz to (my dynamicMenu)
        set indexy to menuitemz's indexOfSelectedItem()
        set the_action to ((my dynamicMenu's titleOfSelectedItem) as Unicode text)
        set datalinee to my menutocellnum(the_action)
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
                    set thrdh to value of (cell ("L" & datalinee))
                    set allauf to my theSplit(thrdh, "*")
                    set theranz to "C" & datalinee
                    set FolderCopy to value of cell theranz
                end tell
            end tell
        end tell
        tell application "Finder"
        repeat with audiofiz in allauf
            if audiofiz contains ":" then
                set audiofiz to audiofiz as alias
                set thef to posix path of audiofiz
                #display dialog quoted form of thef as string
                try
                set commandydoo to do shell script "/usr/local/Cellar/tag/0.10_1/bin/tag -s red " & quoted form of thef
                end try
                move file audiofiz to FolderCopy
            else if audiofiz contains "/" then
                set thef to audiofiz as posix file
                #display dialog thef as string
                try
                set commandydoo to do shell script "/usr/local/Cellar/tag/0.10_1/bin/tag -s red " & "'" & thef & "'"
                end try
                move file thef to FolderCopy
            end if
        end repeat
        end tell
    end Movevotosess_
    on removesess()
        set othersess to false
        set menuitemz to (my dynamicMenu)
        set indexy to menuitemz's indexOfSelectedItem()
        set the_action to ((my dynamicMenu's titleOfSelectedItem) as Unicode text)
        set typemenu to (my dynamicTypeMenu)
        set the_typez to ((typemenu's titleOfSelectedItem) as Unicode text)
        set textybox1 to (my textbox1)
        set cellnumber to my menutocellnum(the_action)
        set tstrvell to textybox1's stringValue as string
        set temp1 to my theSplit(the_action, " | ")
        set tname to (item 1 of temp1 as string)
        set cname to (item 2 of temp1 as string)
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
                    set tempran1 to "I" & cellnumber
                    set value of cell tempran1 to "Approved"
                end tell
            end tell
            try
            tell workbook "Audio Production Sheet.xlsx"
                tell sheet "Audio"
                    set cellnumber2 to ""
                    set tempcelln1 to 3
                    set boolte to false
                    repeat 70 times
                        set tempran1 to "A" & tempcelln1 & ":B" & tempcelln1
                        set celly1 to cell("C" & tempcelln1)
                        set typzzz to value of celly1
                        set blehz to value of range tempran1 as string
                        if blehz contains tname then
                            if blehz contains cname then
                                if (typzzz as string) contains (the_typez as string)
                                set cellnumber2 to tempcelln1
                                set tempran1 to "A" & cellnumber2 & ":H" & cellnumber2
                                set value of range tempran1 to ""
                                set boolte to true
                                #exit repeat
                                else
                                set othersess to true
                                end if
                            end if
                        end if
                        set tempcelln1 to tempcelln1 + 1
                    end repeat
                end tell
            end tell
            end try
        end tell
        if othersess is false then
        menuitemz's removeItemAtIndex_(indexy)
        set listz to menuitemz's itemTitles() as list
        #display dialog listz as string
        tell application "Finder"
            set tempdirz to my psessdir
            set tempdirx to my zsessdir
            delete file tempdirz
            make new file at tempdirx with properties {name:"Previous_sessions.txt", file type:"TEXT", creator type:"ttxt"}
            set open_target_file to (open for access tempdirz with write permission)
            set the_count to count listz
            write "" to open_target_file
            set in_t1 to 1
            repeat the_count times
                set str_z to (item in_t1 of listz) & "*"
                write str_z to open_target_file starting at eof
                set in_t1 to in_t1 + 1
            end repeat
            close access open_target_file
            end tell
        end if
    end removesess
    on removesess2(theMsg)
        set menuitemz to (my dynamicMenu)
        set listz to menuitemz's itemTitles() as list
        tell application "Microsoft Outlook"
            set attachmentsz to attachments of theMsg
            repeat with att1 in attachmentsz
                set savename to name of att1 as string
                if savename contains "mp3" then
                    try
                    set savename to item 1 of my theSplit(savename, "60")
                    end try
                    try
                    set savename to item 1 of my theSplit(savename, "30")
                    end try
                    try
                    set savename to item 1 of my theSplit(savename, "15")
                    end try
                    try
                    set savename to item 1 of my theSplit(savename, "10")
                    end try
                    set thingsz to my getsessiontermff(savename)
                end if
            end repeat
        end tell
        set sessnamee to item 1 of thingsz
        set datalinee to item 2 of thingsz
        repeat with itemz1 in listz
            if itemz1 contains datalinee as string then
                set theRemove to itemz1
                exit repeat
            end if
        end repeat
        set getremove to menuitemz's indexOfItemWithTitle_(theRemove)
        log "removesess2 var theRemove: " & theRemove
        set cellnumber to my menutocellnum(theRemove)
        menuitemz's removeItemAtIndex_(getremove)
        set temp1 to my theSplit(theRemove, " | ")
        set tname to (item 1 of temp1 as string)
        set cname to (item 2 of temp1 as string)
        #display dialog "trying to find in data sheet"
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
                    set tempran1 to "I" & cellnumber
                    set value of cell tempran1 to "Approved"
                end tell
            end tell
            
            try
            tell workbook "Audio Production Sheet.xlsx"
                tell sheet "Audio"
                    set cellnumber2 to ""
                    set tempcelln1 to 3
                    set boolte to false
                    repeat while boolte is false
                        set tempran1 to "A" & tempcelln1 & ":B" & tempcelln1
                        set blehz to value of range tempran1 as string
                        if blehz contains tname then
                            if blehz contains cname then
                                set cellnumber2 to tempcelln1
                                set tempran1 to "A" & cellnumber2 & ":H" & cellnumber2
                                set value of range tempran1 to ""
                                set boolte to true
                                exit repeat
                            end if
                        else if tempcelln1 is greater than 70 then
                            #display dialog "not found in Production sheet"
                        exit repeat
                        end if
                        set tempcelln1 to tempcelln1 + 1
                    end repeat
                end tell
            end tell
            end try
        end tell
        set listz to menuitemz's itemTitles() as list
        tell application "Finder"
            set tempdirz to my psessdir
            set tempdirx to my zsessdir
            delete file tempdirz
            make new file at tempdirx with properties {name:"Previous_sessions.txt", file type:"TEXT", creator type:"ttxt"}
            set open_target_file to (open for access tempdirz with write permission)
            set the_count to count listz
            write "" to open_target_file
            set in_t1 to 1
            repeat the_count times
                set str_z to (item in_t1 of listz) & "*"
                write str_z to open_target_file starting at eof
                set in_t1 to in_t1 + 1
            end repeat
            close access open_target_file
        end tell
    end removesess2
    on looprunner_(sender)
        log "looprunning"
        set my firsttimvp to true
        my performSelectorInBackground_withObject_("runloopscripttp", me)
    end looprunner_
    on setstaus_(sender)
        set the_action to ((my dynamicMenu's titleOfSelectedItem) as Unicode text)
        set textybox1 to (my textbox1)
        set cellnumber to my menutocellnum(the_action)
        set tstrvell to textybox1's stringValue as string
        (*if tstrvell contains ":" as string then
            set newtstringar to my theSplit(tstrvell, " :")
            set newtstring to item 2 of newtstringar
            set tstrvell to item 1 of newtstringar
        else
            set newtstring to "N/A"
        end if
        log "setstatus but var newtstring & tstrvell " & newtstring & tstrvell*)
        set typemenu to (my dynamicTypeMenu)
        set the_typez to ((typemenu's titleOfSelectedItem) as Unicode text)
        set temp1 to my theSplit(the_action, " | ")
        set tname to (item 1 of temp1 as string)
        set cname to (item 2 of temp1 as string)
        my setthestat(tstrvell, cellnumber, the_typez)
    end setstaus_
    on nextstepbut_(sender)
        performSelectorInBackground_withObject_("donextstep", "")
    end nextstepbut_
    on adddocx_(sender)
        set the_action to ((my dynamicMenu's titleOfSelectedItem) as Unicode text)
        set cellnumber to my menutocellnum(the_action)
        
    end adddocx_
    on donextstep()
        set the_action to ((my dynamicMenu's titleOfSelectedItem) as Unicode text)
        set cellnumber to my menutocellnum(the_action)
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
                    set theSatran to (cell ("I" & cellnumber))
                    set theStatus to value of theSatran as string
                    set titran to (cell ("F" & cellnumber))
                    set titn to value of titran as string
                    set cliran to (cell ("D" & cellnumber))
                    set clin to value of cliran as string
                    set foldran to (cell ("C" & cellnumber))
                    set pathb to (value of foldran as string) & "Bounced Files:"
                    set ranzx to "H" & cellnumber
                    set docxH to (value of cell ranzx as string)
                    set typeran to (cell ("G" & cellnumber))
                    set typesG to value of typeran
                    set typeran to (cell ("E" & cellnumber))
                    set clicon to value of typeran
                end tell
            end tell
        end tell
        log theStatus & titn & clin & pathb & docxH
        set sessname to (titn & " " & clin) as string
        if theStatus contains "waiting"
            if theStatus contains "part"
                my bouncedfile(pathb, sessname, cellnumber)
            else if theStatus contains "trans"
            end if
        else if theStatus contains "sent"
            if theStatus contains "comp"
                if theStatus contains "Jerry"
                    my bouncedfile(pathb, sessname, cellnumber)
                else if theStatus contains "Tom"
                try
                    set docxHar to my theSplit(docxH, "*")
                    repeat with doc1 in docxHar
                        if doc1 contains "60" then
                            set docxH to doc1
                            exit repeat
                        else if doc1 contains "TV"
                            if doc1 contains "30"
                                set docxH to doc1
                                exit repeat
                            end if
                        end if
                    end repeat
                end try
                set fold1 to pathb
                set pathbb to "\"" & POSIX path of fold1 & "\""
                set mp3filename to do shell script "cd " & pathbb & " && ls -t *.mp3 | sed -n 1p" as string
                set themp2path to fold1 & mp3filename
                tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/JERRY COMPVO.emltpl" as Posix file)
                my addattachz(themp2path)
                my addattachz(docxH)
                set subbyjj to clin & " COMP VO"
                set bodz to subbyjj
                set subbycjj to subbyjj
                my setthesubabo(subbycjj, bodz)
                my setthestat("COMP VO sent to Jerry", cellnumber, "COMP VO")
                end if
            else if theStatus contains "mix"
                    my approved2(sessname, cellnumber, clin, titn, sendera)
            end if
        end if
    end donextstep
    on startersjp_(sender)
        #performSelectorInBackground_withObject_("scriptjp", "")
        performSelectorInBackground_withObject_("scriptysp", "")
    end startersjp_
    on scriptysp()
        my thefunchtionsj(1)
    end scriptysp
    on scriptjp()
        tell application "Microsoft Outlook"
        set unreadcountsj to (get unread count of mail folder "JP Scripts" of exchange account "jonathan@jerrydefalco.com")
        if unreadcountsj is greater than 1 then
            set listchecked to item 1 through unreadcountsj of (get messages of mail folder "JP Scripts" of exchange account "jonathan@jerrydefalco.com")
            my Newscriptread(listchecked)
        else if unreadcountsj is greater than 0 then
            set listchecked to item 1 of (get messages of mail folder "JP Scripts" of exchange account "jonathan@jerrydefalco.com")
            my Newscriptread(listchecked)
        end if
        end tell
    end scriptjp
    on addsessfold_(sender)
        performSelectorInBackground_withObject_("addsesstocur", "")
    end addsessfold_
    on opensessfilebutwr_(sender)
        performSelectorInBackground_withObject_("opensessfilebut", "")
    end opensessfilebutwr_
    on opendocxbutwr_(sender)
        performSelectorInBackground_withObject_("opendocxbut", "")
    end opendocxbutwr_
    on copydocxbutwr_(sender)
        performSelectorInBackground_withObject_("copydocxbut", "")
    end copydocxbutwr_
    on opensessfolbutwr_(sender)
        performSelectorInBackground_withObject_("opensessfolbut", "")
    end opensessfolbutwr_
    on opendocxbut()
        set the_action to ((my dynamicMenu's titleOfSelectedItem) as Unicode text)
        set typemenu to (my dynamicTypeMenu)
        set the_typez to ((typemenu's titleOfSelectedItem) as Unicode text)
        set cellnumber to my menutocellnum(the_action)
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
                    set tempran1 to "H" & cellnumber
                    set dirptx to value of cell tempran1 as string
                    try
                        set docxHar to my theSplit(dirptx, "*")
                        repeat with doc1 in docxHar
                            if (doc1 as string) contains the_typez as string
                                set dirptx to doc1
                                exit repeat
                            end if
                        end repeat
                    end try
                    tell application "Microsoft Word" to open file (dirptx)
                end tell
            end tell
        end tell
    end opendocxbut
    on copydocxbut()
        set the_action to ((my dynamicMenu's titleOfSelectedItem) as Unicode text)
            set cellnumber to my menutocellnum(the_action)
            set typemenu to (my dynamicTypeMenu)
            set the_typez to ((typemenu's titleOfSelectedItem) as Unicode text)
            tell application "Microsoft Excel"
                tell workbook "Database.xlsx"
                    tell sheet "Main Base"
                        set tempran1 to "H" & cellnumber
                        set dirptx to value of cell tempran1 as string
                        try
                            set docxHar to my theSplit(dirptx, "*")
                            repeat with doc1 in docxHar
                                if (doc1 as string) contains the_typez as string
                                    set dirptx to doc1
                                    exit repeat
                                end if
                            end repeat
                        end try
                        tell application "Keyboard Maestro Engine"
                            setvariable "pathnamez" to dirptx
                            do script "CBFDC3CF-DD28-4292-AFDE-90AAC239E3B9"
                        end tell
                    end tell
                end tell
            end tell
    end copydocxbut
    on addsesstocur()
        set the_action to ((my dynamicMenu's titleOfSelectedItem) as Unicode text)
        set cellnumber to my menutocellnum(the_action)
        tell application "Finder"
        set sessnameA to {}
        set theWin to front Finder window
        set currentfolder to (target of theWin as alias)
        set sessionfiles to every file of (currentfolder as alias)
        set sessionfilesstr to name of every file of (currentfolder as alias) as string
        set numberofsessfiles to count items in sessionfiles
            set file11 to 1
            repeat numberofsessfiles times
                set tempfille to item file11 of sessionfiles as string
                if tempfille contains "ptx" then
                    set end of sessnameA to tempfille
                end if
                set file11 to file11 + 1
            end repeat
                    set sesspathB to item 1 of sessnameA
                    set tempsessname1 to item 1 of sessnameA
                    set tempsessname2 to my theSplit(tempsessname1, ":")
                    set sessnameA to last item of tempsessname2 as string
                    set sessfoldC to currentfolder as alias
                    tell application "System Events" to set sessmoddateN to modification date of file sesspathB
                        end tell
                        tell application "Microsoft Excel"
                            tell workbook "Database.xlsx"
                                tell sheet "Main Base"
                                    set myrange to "A" & cellnumber
                                    set value of cell myrange to sessnameA as string
                                    set myrange to "B" & cellnumber
                                    set value of cell myrange to sesspathB as string
                                    set myrange to "C" & cellnumber
                                    set value of cell myrange to sessfoldC as string
                                    set myrange to "N" & cellnumber
                                    set sessmoddateN to my formatDate(sessmoddateN)
                                    set value of cell myrange to sessmoddateN as string
                                end tell
                            end tell
                        end tell
    end addsesstocur
    on opensessfilebut()
        set the_action to ((my dynamicMenu's titleOfSelectedItem) as Unicode text)
        set cellnumber to my menutocellnum(the_action)
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
                    set tempran1 to "B" & cellnumber
                    set dirptx to value of cell tempran1 as string
                    tell application "Finder" to open alias (dirptx)
                end tell
            end tell
        end tell
    end opensessfilebut
        on opensessfolbut()
            set the_action to ((my dynamicMenu's titleOfSelectedItem) as Unicode text)
            set cellnumber to my menutocellnum(the_action)
            tell application "Microsoft Excel"
                tell workbook "Database.xlsx"
                    tell sheet "Main Base"
                        set tempran1 to "C" & cellnumber
                        set dirptx to value of cell tempran1 as string
                        tell application "Finder" to open alias (dirptx)
                    end tell
                end tell
            end tell
        end opensessfolbut
    on menutocellnum(menustr)
        log "menucellnum" & menustr
        set temp1 to my theSplit(menustr, "<")
        set temp2 to item 2 of temp1
        set temp1 to my theSplit(temp2, ">")
        set cellnumber to item 1 of temp1
        return(cellnumber)
    end menutocellnum
    on addselectoash_(Sender)
        set the_action to ((my dynamicMenu's titleOfSelectedItem) as Unicode text)
        set typemenu to (my dynamicTypeMenu)
        set the_typez to ((typemenu's titleOfSelectedItem) as Unicode text)
        set arrys1 to my theSplit(the_action, " | ")
        set arrys2 to item 2 of my theSplit(the_action, "<")
        set dataline to my theSplit(arrys2, ">") as string
        if (item 3 of arrys1) contains "Jerry" then
            set clicon to "(Jerry)"
        else
            set clicon to "(Jerry)"
        end if
        set cliA to (item 2 of arrys1) & clicon
        set titB to item 1 of arrys1
        set dorepeat to false
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
                    set rangx to "I" & dataline
                    set statusD to value of cell rangx
                    set rangx to "G" & dataline
                    set typeC to value of cell rangx
                    if typeC contains "*"
                    set atypesc to my theSplit(typeC, "*")
                    try
                    repeat with tyypez in atypesc
                        if (tyypez as string) contains the_typez as string
                            set typeC to doc1
                            exit repeat
                        end if
                    end repeat
                    end try
                    end if
                    set rangx to "O" & dataline
                    set SubdE to value of cell rangx
                    set rangx to "J" & dataline
                    set NeedbF to value of cell rangx
                    set rangx to "K" & dataline
                    set Anncr to value of cell rangx
                end tell
            end tell
            set Anncrsh to {}
            if Anncr contains "Mark" then
                if Anncr contains "Mark B" then
                    set end of Anncrsh to "MB "
                else
                    set end of Anncrsh to "MM "
                end if
            end if
                if Anncr contains "Rachel" then
                    set end of Anncrsh to "RB "
                end if
                if Anncr contains "Jim" then
                    set end of Anncrsh to "JM "
                end if
                if Anncr contains "Chris" then
                    set end of Anncrsh to "CC "
                end if
                if Anncr contains "Melissa" then
                    set end of Anncrsh to "MEL "
                end if
                if Anncr contains "Mike O" then
                    set end of Anncrsh to "MO "
                end if
                if Anncr contains "Donovan" then
                    set end of Anncrsh to "DV "
                end if
                if Anncr contains "Brent" then
                    set end of Anncrsh to "BM "
                end if
                if Anncr contains "Andrea" then
                    set end of Anncrsh to "AB "
                end if
                if Anncr contains "Ben" then
                    set end of Anncrsh to "BB "
                end if
                if Anncr contains "David" then
                    set end of Anncrsh to "DT "
                end if
                if Anncr contains "Mitch" then
                    set end of Anncrsh to "MP "
                end if
                if Anncr contains "Paco" then
                    set end of Anncrsh to "PL "
                end if
                if Anncr contains "Sandra" then
                    set end of Anncrsh to "SS "
                end if
                if Anncr contains "Doak" then
                    set end of Anncrsh to "DB "
                end if
                if Anncr contains "Rob" then
                    set end of Anncrsh to "RM "
                end if
            tell workbook "Audio Production Sheet.xlsx"
                tell sheet "Audio"
                    set int2 to 3
                    repeat 50 times
                        set ran3 to "B" & int2
                        set tsessname to value of cell ran3
                        if tsessname is "" then
                            set ndata to int2
                            exit repeat
                        end if
                        set int2 to int2 + 1
                    end repeat
                    set ran3 to "A" & ndata
                    set value of cell ran3 to cliA
                    set ran3 to "B" & ndata
                    set value of cell ran3 to titB
                    set ran3 to "C" & ndata
                    set value of cell ran3 to typeC
                    set ran3 to "D" & ndata
                    set value of cell ran3 to statusD
                    set ran3 to "E" & ndata
                    set value of cell ran3 to SubdE
                    set ran3 to "F" & ndata
                    set value of cell ran3 to NeedbF
                    set ran3 to "G" & ndata
                    set value of cell ran3 to Anncrsh as string
                end tell
            end tell
        end tell
    end addselectoash_
    on idle
        log "Iddddddddleeeeee"
       return 9999999999999
    end idle
on runloopscripttp()
    set myzss to (my switch1)'s state
    if (myzss as string) is "1" then
    #try
    my runloopscriptt()
    #on error
    #delay 15
        #my runloopscriptt()
        #end try
    end if
end runloopscripttp
on runloopscriptt()
    
    set tomvacay to false
    set jvacay to false
    set domvacay to false
    set firsttimv to my firsttimvp
            set lastfivesp to my lastfivespp
            set lastfiveij to my lastfiveijp
            set lastfivesj to my lastfivesjp
            set lastfivetvj to my lastfivetvjp
            set lastsessname to my lastsessnamep
            set sessnameszla to my sessnameszlap
            set datalinef to my datalinefp
            set mesgfivesp to {}
            set mesgfiveij to {}
            set mesgfivesj to {}
            set mesgfivetvj to {}
            set sessfilyz to {}
            set sessnamesz to ""
            set sessionname to ""
            tell application "Microsoft Outlook"
                with timeout of 3600 * 24 seconds
            if firsttimv is true
                set unreadcountsp to (get unread count of inbox of exchange account "spots@jerrydefalco.com")
                #set unreadcountij to (get unread count of inbox of exchange account "jonathan@jerrydefalco.com")
                set unreadcountsj to (get unread count of mail folder "JP Scripts" of exchange account "jonathan@jerrydefalco.com")
                #set unreadcounttvj to (get unread count of mail folder "TV Audio" of exchange account "jonathan@jerrydefalco.com")
                
                (*if unreadcountij is greater than 1 then
                    set listchecked to item 1 through unreadcountij of (get messages of inbox of exchange account "jonathan@jerrydefalco.com")
                    my thefunchtionij(listchecked)
                else if unreadcountij is greater than 0 then
                    set listchecked to item 1 of (get messages of inbox of exchange account "jonathan@jerrydefalco.com")
                    my thefunchtionij(listchecked)
                end if*)
                if unreadcountsj is greater than 1 then
                    set listchecked to item 1 through unreadcountsj of (get messages of mail folder "JP Scripts" of exchange account "jonathan@jerrydefalco.com")
                    my Newscriptread(listchecked)
                else if unreadcountsj is greater than 0 then
                    set listchecked to item 1 of (get messages of mail folder "JP Scripts" of exchange account "jonathan@jerrydefalco.com")
                    my Newscriptread(listchecked)
                end if
                if unreadcountsp is greater than 1 then
                    set listchecked to item 1 through unreadcountsp of (get messages of inbox of exchange account "spots@jerrydefalco.com")
                    my thefunchtionisp(listchecked)
                else if unreadcountsp is greater than 0 then
                    set listchecked to item 1 of (get messages of inbox of exchange account "spots@jerrydefalco.com")
                    my thefunchtionisp(listchecked)
                end if
                (*if unreadcounttvj is greater than 1 then
                    set listchecked to item 1 through unreadcounttvj of (get messages of mail folder "TV Audio" of exchange account "jonathan@jerrydefalco.com")
                    my thefunchtiontvj(listchecked)
                else if unreadcounttvj is greater than 0 then
                    set listchecked to item 1 of (get messages of mail folder "TV Audio" of exchange account "jonathan@jerrydefalco.com")
                    #my thefunchtiontvj(listchecked)
                    display dialog "New TV Audio"
                end if*)
                set mesgfivesp to (get every message of inbox of exchange account "spots@jerrydefalco.com")
                #set mesgfiveij to item 1 through 5 of (get messages of inbox of exchange account "jonathan@jerrydefalco.com")
                set mesgfivesj to item 1 through 5 of (get messages of mail folder "JP Scripts" of exchange account "jonathan@jerrydefalco.com")
                #set mesgfivetvj to item 1 through 5 of (get messages of mail folder "TV Audio" of exchange account "jonathan@jerrydefalco.com")
           else
                set mesgfivesp to (get every message of inbox of exchange account "spots@jerrydefalco.com")
                #set mesgfiveij to item 1 through 5 of (get messages of inbox of exchange account "jonathan@jerrydefalco.com")
                set mesgfivesj to item 1 through 5 of (get messages of mail folder "JP Scripts" of exchange account "jonathan@jerrydefalco.com")
                #set mesgfivetvj to item 1 through 5 of (get messages of mail folder "TV Audio" of exchange account "jonathan@jerrydefalco.com")
            end if
                set sessionnamear to my getopensessname()
                set sessionname to item 1 of sessionnamear
                set nosess to item 2 of sessionnamear
                tell application "Microsoft Excel"
                if nosess is false then
                    tell workbook "Database.xlsx"
                        tell sheet "Main Base"
                    if sessionname is not lastsessname then
                            set myRangez to range "A:A"
                            set cellznum to find myRangez what sessionname
                            set rangz1 to (first row index of cellznum)
                            set datalinef to rangz1
                    end if
                        set ranzx to "C" & datalinef
                        set filelocaf to ((value of cell ranzx as string) & "Bounced Files:")
                    #display dialog filelocaf & "Looper func"
                    try
                        tell application "Finder"
                            set sessfilyz to get (every file of folder filelocaf)
                        end tell
                        if firsttimv is true then
                            set lastsessname to sessionname
                            set sessnameszla to sessfilyz
                        else
                        set countn to count sessfilyz
                        set counto to count sessnameszla
                        #display dialog (sessionname as string) & firsttimv
                        if lastsessname is not "N/A"
                            if sessnameszla is not sessfilyz then
                                if sessionname is lastsessname then
                                    if countn is greater than counto then
                                    my bouncedfile(filelocaf, sessionname, datalinef)
                                    set sessnameszla to sessfilyz
                                    end if
                                end if
                                    
                            end if
                        end if
                        end if
                    end try
                    end tell
                    end tell
                    set lastsessname to sessionname as string
                else
                set lastsessname to "N/A"
                end if
                try
                    set sessnameszla to sessfilyz
                end try
            end tell
            if firsttimv is true then
                set lastfivesp to mesgfivesp
                #set lastfiveij to mesgfiveij
                set lastfivesj to mesgfivesj
                #set lastfivetvj to mesgfivetvj
            end if
            
            (*if mesgfiveij is not lastfiveij then
                set listchecked to my checkdiflist(mesgfiveij, lastfiveij)
                #display dialog "new stuffij" & subject of item 1 of listchecked as string
                
                my thefunchtionij(listchecked)
                
            end if*)
            if mesgfivesj is not lastfivesj then
                set listchecked to my checkdiflist(mesgfivesj, lastfivesj)
                #display dialog "new stuffsj"
                my Newscriptread(listchecked)
            end if
            if mesgfivesp is not lastfivesp then
            if (count of lastfivesp) is greater than (count of mesgfivesp) then
                else
                set listchecked to my checkdiflist(mesgfivesp, lastfivesp)
                #display dialog "new stuffsp"
                #set continone to true
                my thefunchtionisp(listchecked)
                end if
            end if
            (*if mesgfivetvj is not lastfivetvj then
                set listchecked to my checkdiflist(mesgfivetvj, lastfivetvj)
                #display dialog "new stufftvj"
                #set continone to true
                display dialog "New TV Audio"
                #my thefunchtiontvj(listchecked)
            end if*)
            set lastfivesp to mesgfivesp
            #set lastfiveij to mesgfiveij
            set lastfivesj to mesgfivesj
            #set lastfivetvj to mesgfivetvj
            #delay 15
            set firsttimv to false
            #display dialog "continue?"
            end timeout
            end tell
            log "loopscript var: " & firsttimv as string
            try
                log filelocaf as string
            end try
            delay 7
            log "Open sess name: " & sessionname
            set my firsttimvp to firsttimv
            set my lastfivespp to lastfivesp
            #set my lastfiveijp to lastfiveij
            set my lastfivesjp to lastfivesj
            #set my lastfivetvjp to lastfivetvj
            set my lastsessnamep to lastsessname
            set my sessnameszlap to sessnameszla
            set my datalinefp to datalinef
            try
                my runloopscripttp()
            on error
                delay 15
                my runloopscripttp()
            end try
end runloopscriptt
    on thefunchtionisp(fmesgfives)
        #new email for Spots inbox
        log "thefunchtionisp started"
        tell application "Microsoft Outlook"
            set numinij to count fmesgfives
            set numinijr to 1
            set filesenders to my readFile("/Users/jonathanstoff/Documents/macrostuff/email search/filesenders.txt") as string
            set dropboxsenders to my readFile("/Users/jonathanstoff/Documents/macrostuff/email search/dropboxsenders.txt") as string
            repeat with theMsg in fmesgfives
                #set theMsg to item numinijr of fmesgfives
                set theSender to address of (get theMsg's sender) as string
                set contentz to plain text content of theMsg
                if contentz contains "transl"
                    my translatedfile(theMsg)
                #add script for spain scripts
                #else if theSender contains "brent"
                    #my brentsender(theMsg)
                else if filesenders contains theSender
                    my filesendies(theMsg)
                else if dropboxsenders contains theSender
                    my boxsendies(theMsg)
                else if theSender contains "Ben"
                    my bennysendies(theMsg)
                end if
                set numinjr to numinijr + 1
                set is read of theMsg to true
            end repeat
        end tell
    end thefunchtionisp
    on translatedfile(theMsg)
        log "translatedfile"
        tell application "Microsoft Outlook"
            set subby to subject of theMsg as string
            set outputz to my getsessionnameupfz(theMsg)
            set NstoreFolder to "LaCie:current scripts :completed scripts" as alias
            set selection to inbox of exchange account "spots@jerrydefalco.com"
            delay 3
            set selection to theMsg
            delay 3
            set msg to first item of (get current messages)
                set storeFolder to "Macintosh HD:Users:jonathanstoff:Downloads" as alias
                set allAttachments to attachments of msg
                repeat with thisAttachment in allAttachments
                    set saveName to name of thisAttachment
                    save thisAttachment in storeFolder
                end repeat
        end tell
        tell application "Finder"
            set docxHH to get every item of storeFolder
            set NstoreFolder to "LaCie:current scripts :completed scripts" as alias
            set docxH to {}
        repeat with docxdir in docxHH
            set tsn to name of docxdir as string
            if tsn contains ".doc" then
                if tsn contains "AB" then
                set end of docxH to ((NstoreFolder as string) & tsn)
                move docxdir to NstoreFolder with replacing
                delete docxdir
                else
                delete docxdir
                end if
            else
                delete docxdir
            end if
        end repeat
        end tell
        repeat with theDocxf in docxH
            set sessnameA to ""
            set clientnD to "" #yes
            set clientcE to "" #yes
            set titlenF to "" #yes
            set typesG to {} #yes
            set docxpathH to "" #yes
            set vaK to {} #yes
            set clientname to ""
            set titlename to ""
            set newcellnum to 0
            
            tell application "Finder" to open theDocxf
                delay 3
            tell application "Microsoft Word"
                set theDoc to name of active document as string
                set theWin to name of front window
                set docxpathH to theDocxf
                set theSelectionn to selection
                tell selection
                    delay 1
                    set selection start to 2
                    set selection end to 2000
                    delay 2
                        set document1 to content as string
                            
                            set lines1 to my theSplit(document1, "TITLE:")
                            set line1 to item 1 of lines1
                            set line2 to item 2 of lines1
                            if line2 contains "NEED" then
                                set lines2 to my theSplit(line2, "NEED")
                            else if line2 contains "AIR" then
                                set lines2 to my theSplit(line2, "AIR")
                            end if
                            set line2 to item 1 of lines2
                            set line3 to item 2 of lines2
                            set lines3 to my theSplit(line3, "MUSIC:")
                            set line3 to item 1 of lines3
                            set line1array to my theSplit(line1, "spot:")
                            set clientname2 to item 1 of line1array as text
                            set spottype to item 2 of line1array as text
                            set clientname to item 1 of line1array as text
                            set line2array to line2 as string
                            set titlename to line2array
                            if titlename contains "REV" then
                                set fulltit to titlename
                                set titlenamedd to my theSplit(titlename, "REV")
                                set titlename to item 1 of titlenamedd
                                set titlename to my trimThis(titlename, true, "full")
                                set revision1 to true
                            else
                                set fulltit to titlename
                                set revision1 to false
                            end if
                            try
                                if titlename contains "’" then
                                    set titlename to my theSplit(titlename, "’") as string
                                    set fulltit to my theSplit(fulltit, "’") as string
                                end if
                            end try
                            set titlename to my trimThis(titlename, true, "full")
                            set fulltit to my trimThis(fulltit, true, "full")
                            set clientname to my trimThis(clientname, true, "full")
                            set spottype to my trimThis(spottype, true, "full")
                            if clientname contains ":" then
                                set clientname1 to my theSplit(clientname, ":")
                                set inttyc to (count of clientname1) - 1
                                set ehh to 2
                                set clientnamezx to {}
                                repeat inttyc times
                                    if item ehh of clientname1 is not "" then
                                        set end of clientnamezx to item ehh of clientname1
                                    end if
                                    
                                end repeat
                                set clientname to clientnamezx as string
                                set clientname to my trimThis(clientname, true, "full")
                            end if
                            
                            set titlenF to fulltit
                            set typesG to spottype as string
                end tell
                        set vaK to {}
                        if document1 contains "Male" then
                            set end of vaK to "Paco "
                        end if
                        if document1 contains "Girl" then
                            set end of vaK to "Andrea "
                        else if document1 contains "female"
                            set end of vaK to "Andrea "
                        else if document1 contains "woman"
                            set end of vaK to "Andrea "
                        end if
                        if line3 contains "WILNETTE" then
                            set end of vaK to "Wilnette"
                       end if
            if (my printer) is true then
               print active document
            end if
                quit
            end tell
                set cellnumber to item 2 of outputz
                delay 1
                tell application "Microsoft Excel"
                    tell workbook "Database.xlsx"
                            tell sheet "Main Base"
                        set myrange to "H" & cellnumber
                        if value of cell myrange is "" then
                            set value of cell myrange to docxpathH as string
                        else
                            set value of cell myrange to value of cell myrange & "*" & docxpathH as string
                        end if
                        if vaK is not {} then
                        set myrange to "K" & cellnumber
                        set value of cell myrange to vaK as string
                        end if
                        set mynerang to cell ("J" & cellnumber)
                        set datedue to value of mynerang
                        end tell
                 end tell
                end tell
                set Spanish1 to false
                set datedue to my formatDate(datedue)
                        my setthestat("Waiting on Parts", cellnumber, "N/A")
                        my sendvoemail(vaK, clientnD, Spanish1, datedue, docxpathH)
                
        end repeat
    end translatedfile
    on bennysendies(theMsg)
        tell application "Microsoft Outlook" to set mescon to content of theMsg as string
        tell application "Microsoft Outlook" to set sendername to name of (get theMsg's sender) as string
           set link1 to "https://www.benvoice.com/clients/defalco"
    if mescon contains link1 then
    tell application "Safari"
        open location link1
        delay 1
        set inputhtml to do JavaScript "document.documentElement.outerHTML;" in document 1
        set inputar to my theSplit(inputhtml, "<div class=\"block\"><a href=\"")
        set filenametemp to item -1 of inputar
        set filename to item 1 of my theSplit(filenametemp, "\" class=\"mp3\">")
        set filename to sendername & filename
        set link2 to link1 & "/" & filename
        quit
    end tell
    do shell script "cd /Volumes/LaCie/temp && curl " & link2 & " -o " & filename
    tell application "Google Chrome"
        set thetab to active tab of front window
        tell thetab to close
    end tell
        set audiofilename to filename
        set audiofiledir to "LaCie:temp:" & filename
        tell application "Microsoft Outlook"
        set msg to theMsg
        set idz to id of theMsg as string

        set bleh to true
        try
        set sessdat to my getsessionnameupfz(theMsg)
        set sessnamee to item 1 of sessdat
        set datalinee to item 2 of sessdat
        if datalinee is 0 then
            set bleh to false
        end if
        on error
        set bleh to false
        end try
        if bleh is false
            set namez to audiofilename
            set dir1 to audiofiledir
            set thingsz to my getsessiontermff(namez)
            set sessnamee to item 1 of thingsz
            set datalinee to item 2 of thingsz
            my theChekening(datalinee, dir1, sendername, namez, idz, theMsg, sessnamee)
        else
            set namez to audiofilename
            set dir1 to audiofiledir
            my theChekening(datalinee, dir1, sendername, namez, idz, theMsg, sessnamee)
        end if
        end tell
        end if
    end bennysendies
    on filesendies(theMsg)
    log "filesendies started"
        set audiofilename to {}
        set audiofiledir to {}
        set audiofilenamez to {}
        tell application "Microsoft Outlook"
        set msg to theMsg
        set idz to id of theMsg as string
        set sendername to name of (get theMsg's sender) as string
        set bleh to true
        try
        set sessdat to my getsessionnameupfz(theMsg)
        set sessnamee to item 1 of sessdat
        set datalinee to item 2 of sessdat
        if datalinee is 0 then
            set bleh to false
        end if
        on error
        set bleh to false
        end try
        set sandboxDocumentFolder to (path to documents folder as string)
        set sandboxAttachmentsFolder to sandboxDocumentFolder & "Attachments"
        set storeFolder to "LaCie:temp" as alias
        #set storeFolder to "Macintosh HD:Users:jonathanstoff:temp" as alias
        set allAttachments to attachments of msg
        set attcon to count of allAttachments
        repeat with thisAttachment in allAttachments
            set saveName to name of thisAttachment
            set tsaveName to (sendername & " " & saveName) as string
            set saveName1 to my filzexists(storeFolder, tsaveName)
            if saveName contains "mp3"
                set end of audiofilenamez to saveName
                save thisAttachment in storeFolder
                set end of audiofiledir to storeFolder & saveName1
                set end of audiofilename to saveName1
                tell application "Finder" to set name of (alias ((storeFolder & saveName) as string)) to saveName1 as string
            else if saveName contains "wav"
                set end of audiofilenamez to saveName
                save thisAttachment in storeFolder
                set end of audiofiledir to storeFolder & saveName1
                set end of audiofilename to saveName1
                tell application "Finder" to set name of (alias ((storeFolder & saveName) as string)) to saveName1 as string
            else
                set attcon to attcon - 1
            end if
            end repeat
        end tell
            log "filesendies downloaded"
        
        if bleh is false then
            log "filesendies var bleh is false"
            if attcon is greater than 1
                log "filesendies var attcon is " & attcon as string
                set int1 to 1
                repeat attcon times
                    log "filesendies var attcon is " & attcon as string
                        set namez to item int1 of audiofilenamez
                        set dir1 to item int1 of audiofiledir
                        if sendername contains "Mark" then
                            set namez to my theSplit(namez, "- ") as string
                            else if sendername contains "mel" as string then
                                set namez to item 1 of my theSplit(namez, "_")
                            else if sendername contains "rachel" as string then
                            try
                                    set namez to item 1 of my theSplit(namez, "-")
                                    set namez to (my theSplit(namez, "mitz") as string) & "mitsubishi"
                            end try
                            else if sendername contains "chris" as string then
                                set namez to item 1 of my theSplit(namez, "(")
                                set namez to item 1 of my theSplit(namez, "6")
                        else
                        end if
                        set thingsz to my getsessiontermff(namez)
                        log "filesendies var thingsz is " & thingsz as string
                        set sessnamee to item 1 of thingsz
                        set datalinee to item 2 of thingsz
                        #log "filesendies" & (datalinee, dir1, sendername, namez, idz, theMsg, sessnamee) as string
                        my theChekening(datalinee, dir1, sendername, namez, idz, theMsg, sessnamee)
                    set int1 to int1 + 1
                end repeat
            else if attcon is 1
            set namez to audiofilenamez as string
            log "filesendies var: " & namez
            set dir1 to audiofiledir as string
            if sendername contains "Mark" then
                #log "filesendies var sendername: " & sendername
                set namez to my theSplit(namez, "- ") as string
            else if sendername contains "mel" as string then
                set namez to item 1 of my theSplit(namez, "_")
            else if sendername contains "rachel" as string then
            try

                set namez to item 1 of my theSplit(namez, "-")

                if namez contains "mitz" as string then
                    set namez to (my theSplit(namez, "mitz") as string) & "mitsubishi"
                end if
            end try
            else if sendername contains "sand" as string then
                if namez contains "sosub" then
                    set namez to "South Suburban " & namez as string
                end if
            end if
            set thingsz to my getsessiontermff(namez)
            set sessnamee to item 1 of thingsz
            set datalinee to item 2 of thingsz
            #log "filesendies" & (datalinee, dir1, sendername, namez, idz, theMsg, sessnamee) as string
            my theChekening(datalinee, dir1, sendername, namez, idz, theMsg, sessnamee)
            end if
        else
        log "filesendies var bleh is true"
        if attcon is greater than 1
            log "filesendies var attcon is " & attcon as string
            set int1 to 1
            repeat attcon times
                    set namez to item int1 of audiofilenamez
                    set dir1 to item int1 of audiofiledir
                    set thingsz to my getsessiontermff(namez)
                    set sessnamee to item 1 of thingsz
                    set datalinee to item 2 of thingsz
                    #log "filesendies" & (datalinee, dir1, sendername, namez, idz, theMsg, sessnamee) as string
                    my theChekening(datalinee, dir1, sendername, namez, idz, theMsg, sessnamee)
                set int1 to int1 + 1
            end repeat
        else if attcon is 1
        set namez to audiofilenamez as string
        set dir1 to audiofiledir as string
        #og "filesendies" & (datalinee, dir1, sendername, namez, idz, theMsg, sessnamee) as string
        my theChekening(datalinee, dir1, sendername, namez, idz, theMsg, sessnamee)
        end if
        end if
    end filesendies
        on boxsendies(theMsg)
            set audiofilename to ""
            set audiofiledir to ""
            tell application "Microsoft Outlook"
            set msg to theMsg
            set msgcon to content of theMsg
            set idz to id of theMsg as string
            set sendername to name of (get theMsg's sender) as string
            set bleh to true
            try
            set sessdat to my getsessionnameupfz(theMsg)
            set sessnamee to item 1 of sessdat
            set datalinee to item 2 of sessdat
            if datalinee is 0 then
                set bleh to false
            end if
            on error
            set bleh to false
            end try
            set msgcon to content of msg
            end tell
                if msgcon contains "class=\"\">https://www.dropbox" then
                set mest1 to item 2 of my theSplit(msgcon, "class=\"\">https://www.dropbox")
                set mest2 to item 1 of my theSplit(mest1, "</a>")
                try
                set mest3 to item 1 of my theSplit(mest2, "dl=0")
                end try
                set link1 to "https://www.dropbox" & (mest3 as string) & "dl=1"
            tell application "Safari"
                open location link1
            end tell
            tell application "Finder"
                set pathtodownloads to ("/Users/jonathanstoff/Downloads" as POSIX file)
                set bleh to true
                repeat while bleh is true
                    try
                        repeat while bleh is true
                if link1 contains ".wav"
                    set audiofiledir to item 1 of (every item of container pathtodownloads of application "Finder" whose name contains ".wav")
                else if link1 contains ".mp3"
                    set audiofiledir to item 1 of (every item of container pathtodownloads of application "Finder" whose name contains ".mp3")
                end if
                if name of audiofiledir contains "download"
                    
                else
                    exit repeat
                end if
                    end repeat
                set namez to name of audiofiledir as string
                set audiofilename to name of audiofiledir
                set audiofilename to sendername & audiofilename
                set name of audiofiledir to audiofilename
                set audiofiledir to "Macintosh HD:Users:jonathanstoff:Downloads:" & audiofilename
                set storeFolder to "LaCie:temp" as alias
                move audiofiledir to storeFolder with replacing
                delete audiofiledir
                set audiofiledir to storeFolder & audiofilename
                exit repeat
                on error
                end try
            end repeat
                tell application "Safari" to quit
                if bleh is false then
                    #set namez to audiofilename as string
                    set dir1 to audiofiledir as string
                    set thingsz to my getsessiontermff(namez)
                    set sessnamee to item 1 of thingsz
                    set datalinee to item 2 of thingsz
                    my theChekening(datalinee, dir1, sendername, namez, idz, theMsg, sessnamee)
                else
                #set namez to audiofilename as string
                set dir1 to audiofiledir as string
                my theChekening(datalinee, dir1, sendername, namez, idz, theMsg, sessnamee)
                end if
                
            end tell
            end if
        end boxsendies
                on brentsender(theMsg)
                    set datalinee to 0
                    set audiofilename to ""
                    set audiofiledir to ""
                    tell application "Microsoft Outlook"
                    set msg to theMsg
                    set msgcon to content of theMsg
                    set idz to id of theMsg as string
                    set sendername to name of (get theMsg's sender) as string
                    set bleh to true
                    try
                    set sessdat to my getsessionnameupfz(theMsg)
                    set sessnamee to item 1 of sessdat
                    set datalinee to item 2 of sessdat
                    if datalinee is 0 then
                        set bleh to false
                    end if
                    on error
                    set bleh to false
                    end try
                    set msgcon to content of msg
                    #display dialog msgcon as string
                    end tell
                        if msgcon contains "href=\"https://drive.google.com" then
                        set mest1 to item 2 of my theSplit(msgcon, "href=\"https://drive.google.com/file/d/")
                        set mest2 to item 1 of my theSplit(mest1, "/view?usp=drive_web")
                        #try
                        #set mest3 to item 1 of my theSplit(mest2, "dl=0")
                        #end try
                        set link1 to "https://drive.google.com/uc?id=" & (mest2 as string) & "&export=download"
                    tell application "Safari"
                        open location link1
                    end tell
                    tell application "Finder"
                        set pathtodownloads to ("/Users/jonathanstoff/Downloads" as POSIX file)
                        set blehz to true
                        repeat while blehz is true
                            try
                                repeat while blehz is true
                        if link1 contains ".wav"
                            set audiofiledir to item 1 of (every item of container pathtodownloads of application "Finder" whose name contains ".wav")
                        else if link1 contains ".mp3"
                            set audiofiledir to item 1 of (every item of container pathtodownloads of application "Finder" whose name contains ".mp3")
                        end if
                        if name of audiofiledir contains "download"
                        else
                            exit repeat
                        end if
                            end repeat
                        set namez to name of audiofiledir as string
                        set audiofilename to name of audiofiledir
                        set audiofilename to sendername & audiofilename
                        set name of audiofiledir to audiofilename
                        set audiofiledir to "Macintosh HD:Users:jonathanstoff:Downloads:" & audiofilename
                        set storeFolder to "LaCie:temp" as alias
                        move audiofiledir to storeFolder with replacing
                        delete audiofiledir
                        set audiofiledir to storeFolder & audiofilename
                        exit repeat
                        on error
                        end try
                    end repeat
                        tell application "Safari" to quit
                            if bleh is false then
                                log "filesendies var bleh is false"
                                set namez to audiofilename
                                log "filesendies var: " & namez
                                set dir1 to audiofiledir as string
                                if sendername contains "Mark" then
                                    #log "filesendies var sendername: " & sendername
                                    set namez to item 1 of my theSplit(namez, "- ") as string
                                else if sendername contains "mel" as string then
                                    set namez to item 1 of my theSplit(namez, "_")
                                else if sendername contains "rachel" as string then
                                try
                                        set namez to item 1 of my theSplit(namez, "-")
                                        set namez to (my theSplit(namez, "mitz") as string) & "mitsubishi"
                                        
                                end try
                                end if
                                set thingsz to my getsessiontermff(namez)
                                log "filesendies var thingsz is " & thingsz as string
                                set sessnamee to item 1 of thingsz
                                set datalinee to item 2 of thingsz
                                #log "filesendies" & (datalinee, dir1, sendername, namez, idz, theMsg, sessnamee) as string
                                my theChekening(datalinee, dir1, sendername, namez, idz, theMsg, sessnamee)
                            else
                            
                            log "filesendies var bleh is true"
                            set namez to audiofilename
                            set dir1 to audiofiledir as string
                            #og "filesendies" & (datalinee, dir1, sendername, namez, idz, theMsg, sessnamee) as string
                            my theChekening(datalinee, dir1, sendername, namez, idz, theMsg, sessnamee)
                            end if
                      
                    end tell
                    end if
                end brentsender
    on thefunchtionij(fmesgfiveij)
    log "function: thefunchtionij"
        #new email for My inbox
        tell application "Microsoft Outlook"
            set numinij to count fmesgfiveij
            set numinijr to 1
            #Found terms related to item reference of
            set resultsar to {}
            set textar to {}
            set replymail to {}
            set fullcont to {}
            set fullcont1 to {}
            repeat numinij times
                set end of resultsar to " "
                set end of textar to " "
                set end of replymail to " "
            end repeat
            repeat numinij times
                set subb1y to (subject of (item numinijr of fmesgfiveij)) as string
                if subb1y contains "Re:" then
                    set item numinijr of replymail to true
                    #display dialog subb1y
                end if
                set tempijcon to plain text content of item numinijr of fmesgfiveij as string
                set item numinijr of textar to tempijcon
                set emailsed to tempijcon
                set tempijconu to my change_case(tempijcon)
                
                try
                    set stempijcon to my theSplit(tempijconu, "IPHONE")
                    set tempijconu to item 1 of stempijcon
                    set stempijcon to my theSplit(tempijconu, "WROTE:")
                    set tempijconu to item 1 of stempijcon
                    set stempijcon to my theSplit(tempijconu, "FOLLOW")
                    set tempijconu to item 1 of stempijcon
                end try
                set searchtermsijstr to my readFile("/Users/jonathanstoff/Documents/macrostuff/email search/Ijlookterms.txt") as string
                set searchtermsijstru to my change_case(searchtermsijstr)
                #set searchtermsijstru to my trimThis(searchtermsijstru, true, "full")
                set searchtermsijar to my theSplit(searchtermsijstru, ";")
                set searchtermsijint to (count searchtermsijar) - 1
                set intone to 1
                #log searchtermsijstru
                repeat searchtermsijint times
                    set readfilesearchstr to item intone of searchtermsijar
                    set readfilesearchstr to my trimThis(readfilesearchstr, true, "full")
                    
                    if readfilesearchstr is " " then
                        
                    else if tempijconu contains readfilesearchstr as string then
                        set item numinijr of resultsar to readfilesearchstr
                        exit repeat
                    else
                        set item numinijr of resultsar to "n/a"
                    end if
                    set intone to intone + 1
                    set end of fullcont1 to tempijconu as string
                end repeat
                set numinijr to numinijr + 1
            end repeat
            set numinijr to 1
            repeat numinij times
                set subbyx to (subject of (item numinijr of fmesgfiveij)) as string
                set didrun to false
                set textstrz to item numinijr of textar as string
                set tempysx to item numinijr of resultsar as string
                if numinij is 1 then
                    set themesss to fmesgfiveij
                else
                    set themesss to (item numinijr of fmesgfiveij)
                end if
                #display dialog tempysx as string
                set getasender to true
                set inty to 1
                set failex to false
                repeat while getasender is true
                try
                if inty is greater than 200
                    set failex to true
                    exit repeat
                end if
                tell application "Microsoft Outlook" to set sendera to address of (get themesss's sender) as string
                set getasender to false
                exit repeat
                on error
                set inty to inty + 1
                set getasender to true
                delay 1
                end try
                end repeat
                if failex is true
                    exit repeat
                end if
                set therepaly to item numinijr of replymail as string
                set thenotmenu to true
                try
                set sessdat to my getsessionnameup(themesss)
                set sessnamee to item 1 of sessdat
                set datalinee to item 2 of sessdat
                
                set thingsz to my getsessiontermff(subbyx)
                set sessnamee to item 1 of thingsz
                set datalinee to item 2 of thingsz
                
                if datalinee is 0 then
                    set thenotmenu to false
                end if
                end try
                #display dialog thenotmenu as string
                if thenotmenu is true
                if tempysx contains "n/a" then
                    
                else if tempysx contains "THEN MIX" then
                    if sendera contains "Jerry@" then
                        if textstrz contains "(%%211comp!!%11234vo%)" then
                        my thenMIX(numinijr, emailsed, themesss)
                        set didrun to true
                        end if
                    else if sendera contains "Max@" then
                        if textstrz contains "(%%191comp!!%11234vo%)" then
                            my thenMIX(numinijr, emailsed, themesss)
                            set didrun to true
                        end if
                    end if
                else if tempysx contains "MIX" then
                #display dialog textstrz & tempysx & address of (get themesss's sender)
                    if sendera contains "Jerry@" then
                        if textstrz contains "(%%211comp!!%11234vo%)" then
                            my Mixnow(numinijr, themesss)
                            set didrun to true
                        end if
                    else if sendera contains "Max@" then
                    #display dialog textstrz & tempysx
                        if textstrz contains "(%%191comp!!%11234vo%)" then
                            my Mixnow(numinijr, themesss)
                            set didrun to true
                        end if
                    end if
                else if tempysx contains "THEN APPROVED" then
                    if sendera contains "Jerry@" then
                        if textstrz contains "(%%16341mix!!%11234%)" then
                            my MIXedits(numinijr, emailsed, themesss)
                            set didrun to true
                        else if textstrz contains "(%%211comp!!%11234vo%)" then
                            my Mixnow(numinijr, themesss)
                            set didrun to true
                        end if
                    else if sendera contains "Max@" then
                        if textstrz contains "(%%16maxy41mix!!%11234%)" then
                            my MIXedits(numinijr, emailsed, themesss)
                            my theReplyz(themesss, subbyx)
                            set didrun to true
                        else if textstrz contains "(%%191comp!!%11234vo%)" then
                            my Mixnow(numinijr, themesss)
                            set didrun to true
                        end if
                    end if
                else if tempysx contains "ROLL" then
                    if sendera contains "Jerry@" then
                        if textstrz contains "(%%16341mix!!%11234%)" then
                            my approved(numinijr, themesss, sendera)
                            set didrun to true
                        end if
                    else if sendera contains "Max@" then
                            my approved(numinijr, themesss, sendera)
                            set didrun to true
                    else if sendera contains "Audio@" then
                                my approved(numinijr, themesss, sendera)
                                set didrun to true
                    end if
                else if tempysx contains "APPROVED" then
                    if sendera contains "Jerry@" then
                        if textstrz contains "(%%16341mix!!%11234%)" then
                            my approved(numinijr, themesss, sendera)
                            set didrun to true
                        else if textstrz contains "(%%211comp!!%11234vo%)" then
                            my Mixnow(numinijr, themesss)
                            set didrun to true
                            my theReplyz(themesss, subbyx)
                        end if
                    else if sendera contains "Max@" then
                        if textstrz contains "(%%16maxy41mix!!%11234%)" then
                            my approved(numinijr, themesss, sendera)
                            set didrun to true
                        end if
                    else if sendera contains "Audio@" then
                        if textstrz contains "attached" then
                            my removesess2(themesss)
                            set didrun to true
                        end if
                    end if
                else if tempysx contains "CUT" then
                        if sendera contains "Tom@" then
                            if textstrz contains "(%%111comp!!%11234vo%)" then
                            my Compvosugcut(themesss)
                            set didrun to true
                            end if
                        else if sendera contains "Jerry@" then
                            if textstrz contains "(%%211comp!!%11234vo%)" then
                                my makecuts(themesss)
                                my theReplyz(themesss, subbyx)
                                set didrun to true
                            end if
                        else if sendera contains "Max@" then
                            if textstrz contains "(%%191comp!!%11234vo%)" then
                                my makecuts(themesss)
                                my theReplyz(themesss, subbyx)
                                set didrun to true
                            end if
                        else if sendera contains "Audio@" then
                            if textstrz contains "(%%111comp!!%11234vo%)" then
                                my makecuts(themesss)
                                my theReplyz(themesss, subbyx)
                                set didrun to true
                                
                            end if
                        end if
                else if tempysx contains "GOOD" then
                #display dialog textstrz
                    if textstrz contains "(%%111comp!!%11234vo%)" then
                        if sendera contains "Tom@" then
                        set temppr to {}
                        set drrayd to every recipients of themesss
                        set doita to count drrayd
                        set doitr to 1
                        repeat doita times
                            set thiscar to item doitr of drrayd
                            set end of temppr to get thiscar's email address
                            set doitr to doitr + 1
                        end repeat
                            if temppr contains "Jerry@" as string then
                            else if temppr contains "Max@" as string then
                            else
                                my Compvogood(themesss)
                                set didrun to true
                            end if
                        end if
                    end if
                end if
                if therepaly is "true" then
                    if didrun is false then
                    if textstrz contains "(%%16341mix!!%11234%)" then
                        my theReplyz(themesss, subbyx)
                        my MIXedits(numinijr, emailsed, themesss)
                    else if textstrz contains "(%%16maxy41mix!!%11234%)" then
                        my theReplyz(themesss, subbyx)
                        my MIXedits(numinijr, emailsed, themesss)
                    else if textstrz contains "(%%191mix!!%11234%)" then
                        my theReplyz(themesss, subbyx)
                        my MIXedits(numinijr, emailsed, themesss)
                    else if textstrz contains "(%%111comp!!%11234vo%)"
                        my theReplyz(themesss, subbyx)
                    else if textstrz contains "(%%211comp!!%11234vo%)"
                        my theReplyz(themesss, subbyx)
                    else if textstrz contains "(%%191comp!!%11234vo%)"
                        my theReplyz(themesss, subbyx)
                    end if
                    end if
                end if
                end if
                set numinijr to numinijr + 1
                try
                    set end of fullcont to textstrz as string
                tell application "Microsoft Outlook" to set is read of themesss to true
                end try
            end repeat
        end tell
            if failex is true
                log "failed to get sender in thefunchtionij"
            end if
            log resultsar as string
            log fullcont as string
            log "Searched email contents:" & fullcont as string
            log "function end thefunchtionij"
    end thefunchtionij
        on MIXedits(intt, stremail, fmesgfiveij)
        set thenotmenu to true
            with timeout of 86400 seconds
                set boolent to true
                repeat while boolent is true
                    tell application "Microsoft Outlook" to display dialog "Mix edits: " & subject of fmesgfiveij as string buttons {"Open session", "Wait 10 mins", "Got it!"}
                    if the button returned of the result is "Open session" then
                        set boolent to false
                        exit repeat
                    else if the button returned of the result is "Wait 10 mins" then
                        delay 600
                    else
                        set boolent to true
                        exit repeat
                    end if
                end repeat
            end timeout
            
            tell application "Microsoft Outlook"
            set subby to subject of fmesgfiveij as string
            set sessdat to my getsessionnameup(fmesgfiveij)
            set sessnamee to item 1 of sessdat
            set datalinee to item 2 of sessdat
            if datalinee is 0 then
                set thenotmenu to false
            end if
            if thenotmenu is false then
            set sessionnamear to my getopensessname()
            set opensessnamee to item 1 of sessionnamear as string
            set nosess to item 2 of sessionnamear
            end if
            end tell
            if thenotmenu is false then
            set subby to my theSplit(subby, ":") as string
            if boolent is false
            if opensessnamee contains sessnamee then
                try
                tell application "Finder"
                    set subby to subby & "_MIX_Edits" & ".txt"
                        set filenameeq to "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" & subby
                        try
                        delete filenameeq
                        end try
                        make new file at "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" with properties {name:subby, file type:"TEXT", creator type:"ttxt"}
                        set tFile to open for access filenameeq with write permission
                        try
                           write (stremail as string) to tFile starting at eof
                           close access tFile
                        on error
                           close access tFile
                        end try
                        open filenameeq
                    set filenameeq to "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" & subby
                    open filenameeq
                end tell
                end try
            else
                opensessfromline(datalinee)
                delay 5
                tell application "Finder"
                    set subby to subby & "_MIX_Edits" & ".txt"
                    
                        set filenameeq to "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" & subby
                        try
                        delete filenameeq
                        end try
                        make new file at "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" with properties {name:subby, file type:"TEXT", creator type:"ttxt"}
                        set tFile to open for access filenameeq with write permission
                        try
                           write (stremail as string) to tFile starting at eof
                           close access tFile
                        on error
                           close access tFile
                        end try
                        open filenameeq
                    set filenameeq to "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" & subby
                    open filenameeq
                end tell
            end if
            tell application "Microsoft Excel"
                tell workbook "Database.xlsx"
                    tell sheet "Main Base"
                        set rangz to "I" & datalinee
                        set value of cell rangz to "mixing"
                    end tell
                end tell
                tell workbook "Audio Production Sheet.xlsx"
                    tell sheet "Audio"
                        set rangz to "D" & datalinee
                        set value of cell rangz to "mixing"
                    end tell
                end tell
            end tell
            end if
            else
            tell application "Finder"
                set subby to subby & "_MIX_Edits" & ".txt"
                    set filenameeq to "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" & subby
                    try
                    delete filenameeq
                    end try
                    make new file at "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" with properties {name:subby, file type:"TEXT", creator type:"ttxt"}
                    set tFile to open for access filenameeq with write permission
                    try
                       write (stremail as string) to tFile starting at eof
                       close access tFile
                    on error
                       close access tFile
                    end try
                    open filenameeq
                set filenameeq to "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" & subby
                open filenameeq
            end tell
            end if
        end MIXedits
    on theReplyz(themesss, subbyx)
    try
        log "func theReplyz"
        tell application "Microsoft Outlook"
            set thecontz to plain text content of themesss as string
            
            set array1 to my getsessionnameupz(themesss, subbyx)
            set datalinee to item 2 of array1
            set idz to id of themesss as string
            if subbyx contains "radio"
                if subbyx contains "60"
                    set idz to "*radio 60:" & idz
                else if subbyx contains "30"
                    set idz to "*radio 30:" & idz
                else
                    set idz to "*radio:" & idz
                end if
            else if subbyx contains "TV"
                if subbyx contains "30"
                    set idz to "*TV 30:" & idz
                else if subbyx contains "15"
                    set idz to "*TV 15:" & idz
                else if subbyx contains "10"
                    set idz to "*TV 10:" & idz
                else
                    set idz to "*TV:" & idz
                end if
            else if subbyx contains "COMP VO"
            if thecontz contains "(%%16341mix!!%11234%)" as string then
                set idz to "*radio:" & idz
            end if
                set idz to "*COMP VO:" & idz
            end if
 
        end tell
        log "func theReplyz: " & idz
    tell application "Microsoft Excel"
        tell workbook "Database.xlsx"
            tell sheet "Main Base"
                set ranz1 to "M" & datalinee
                set value of cell ranz1 to idz
            end tell
        end tell
    end tell
    log "end theReplyz"
    end try
    end theReplyz
    on thefunchtionsj(int1)
        #new email for JP Scripts
        ignoring application responses
            tell application "Keyboard Maestro Engine"
                do script "2 Get latest script"
            end tell
        end ignoring
            tell application "Keyboard Maestro Engine"
                delay 10
                set besz to false
                repeat while besz is false
                set additemnew to getvariable "additem"
                if additemnew is not "" then
                    exit repeat
                    set besz to true
                    end if
                delay 5
                end repeat
            end tell
            my writetopsess(additemnew)
    end thefunchtionsj
    
    on thenMIX(intt, stremail, fmesgfiveij)
        with timeout of 86400 seconds
            set boolent to true
            repeat while boolent is true
                tell application "Microsoft Outlook" to display dialog "Mix approval after edits: " & subject of fmesgfiveij as string buttons {"Open session", "Wait 10 mins", "Got it!"}
                if the button returned of the result is "Open session" then
                    set boolent to false
                    exit repeat
                else if the button returned of the result is "Wait 10 mins" then
                    delay 600
                else
                    set boolent to true
                    exit repeat
                end if
            end repeat
        end timeout
        if boolent is false
            
        tell application "Microsoft Outlook"
        set subby to subject of (fmesgfiveij) as string
        set sessdat to my getsessionnameup(fmesgfiveij)
        set sessnamee to item 1 of sessdat
        set datalinee to item 2 of sessdat
        set sessionnamear to my getopensessname()
        set opensessnamee to item 1 of sessionnamear as string
        set nosess to item 2 of sessionnamear
        end tell
        set subby to my theSplit(subby, ":") as string
        if opensessnamee contains sessnamee then
            tell application "Finder"
                set subby to subby & "_Edits" & ".txt"
                
                    set filenameeq to "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" & subby
                    delete filenameeq
                    make new file at "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" with properties {name:subby, file type:"TEXT", creator type:"ttxt"}
                    set tFile to open for access filenameeq with write permission
                    try
                       write (stremail as string) to tFile starting at eof
                       close access tFile
                    on error
                       close access tFile
                    end try
                    open filenameeq
                set filenameeq to "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" & subby
                open filenameeq
            end tell
        else
            opensessfromline(datalinee)
            delay 5
            tell application "Finder"
                set subby to subby & "_Edits" & ".txt"
                
                    set filenameeq to "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" & subby
                    delete filenameeq
                    make new file at "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" with properties {name:subby, file type:"TEXT", creator type:"ttxt"}
                    set tFile to open for access filenameeq with write permission
                    try
                       write (stremail as string) to tFile starting at eof
                       close access tFile
                    on error
                       close access tFile
                    end try
                    open filenameeq
                set filenameeq to "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" & subby
                open filenameeq
            end tell
        end if
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
                    set rangz to "I" & datalinee
                    set value of cell rangz to "mixing"
                end tell
            end tell
            tell workbook "Audio Production Sheet.xlsx"
                tell sheet "Audio"
                    set rangz to "D" & datalinee
                    set value of cell rangz to "mixing"
                end tell
            end tell
        end tell
        end if
    end thenMIX
    on Mixnow(intt, fmesgfiveij)
        log intt as string
        tell application "Microsoft Outlook"
            set newestmm to fmesgfiveij
            set subbyjj to subject of newestmm as string
        end tell
        with timeout of 86400 seconds
            set boolent to true
            repeat while boolent is true
                display dialog "Mix approval for: " & subbyjj buttons {"Open session", "Wait 10 mins", "Got it"}
                if the button returned of the result is "Open session" then
                    set boolent to false
                    exit repeat
                else if the button returned of the result is "Wait 10 mins" then
                    delay 600
                else
                    set boolent to true
                    exit repeat
                end if
            end repeat
        end timeout
        set sessdat to my getsessionnameup(fmesgfiveij)
        set sessnamee to item 1 of sessdat
        set datalinee to item 2 of sessdat
        if boolent is false
        set sessionnamear to my getopensessname()
        set opensessnamee to item 1 of sessionnamear as string
        set nosess to item 2 of sessionnamear
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
        set ranyg to "I" & datalinee
        set value of cell ranyg to "Mixing" as string
        set thedaytoday1ar to current date
        set thedaytoday2 to (month of thedaytoday1ar & "/" & day of thedaytoday1ar) as string
        set ranyg to "N" & datalinee
        set value of cell ranyg to thedaytoday2 as string
        end tell
        end tell
        end tell
        if opensessnamee contains sessnamee then
            
        else
            opensessfromline(datalinee)
        end if
        end if
        my setthestat("Mixing", datalinee, "N/A")
    end Mixnow
    on bouncedfile(pathb, sessname, datalinef)
        set replycheck to ""
        log "bouncedfile function"
        set pathbb to "\"" & POSIX path of pathb & "\""
        set mp3filename to do shell script "cd " & pathbb & " && ls -t *.mp3 | sed -n 1p" as string
        set themp2path to (pathb & mp3filename) as string
        try
            set pathbc to "\"" & POSIX path of pathb & "\""
            set aifffilename to do shell script "cd " & pathbc & " && ls -t *.aif | sed -n 1p" as string
            set theaiffpath to (pathb & aifffilename) as string
        end try
        log "boundf foundmp3"
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
            set zzrta to "M" & datalinef
            set replycheck to value of cell zzrta
            set ranyggg to "H" & datalinef
            set docpath to value of range ranyggg as string
            end tell
        end tell
        end tell
            try
                if docpath contains "*" then
                    log "has *"
                set docxHar to my theSplit(docpath, "*")
                if mp3filename contains "Spanish" then
                    if mp3filename contains "60 radio" then
                    repeat with doc1 in docxHar
                        if doc1 contains "60" then
                            if doc1 contains "radio" then
                                if doc1 contains "AB" then
                                set docpath to doc1
                                exit repeat
                                end if
                            end if
                        end if
                    end repeat
                    repeat with doc1 in docxHar
                        if doc1 contains "60" then
                            if doc1 contains "radio" then
                                if doc1 does not contain "AB" then
                                set docpath to doc1
                                exit repeat
                                end if
                            end if
                        end if
                    end repeat
                    else if mp3filename contains "30 radio" then
                    repeat with doc1 in docxHar
                        if doc1 contains "30" then
                            if doc1 contains "radio" then
                                if doc1 contains "AB" then
                            set docpath to doc1
                            exit repeat
                            end if
                            end if
                        end if
                    end repeat
                    repeat with doc1 in docxHar
                        if doc1 contains "30" then
                            if doc1 contains "radio" then
                                if doc1 does not contain "AB" then
                            set docpath to doc1
                            exit repeat
                            end if
                            end if
                        end if
                    end repeat
                    else if mp3filename contains "COMP VO" then
                    repeat with doc1 in docxHar
                        if doc1 contains "60" then
                            if doc1 contains "radio" then
                                if doc1 contains "AB" then
                                set docpath to doc1
                                exit repeat
                                end if
                            end if
                        end if
                    end repeat
                    repeat with doc1 in docxHar
                        if doc1 contains "60" then
                            if doc1 contains "radio" then
                                if doc1 does not contain "AB" then
                                set docpath to doc1
                                exit repeat
                                end if
                            end if
                        end if
                    end repeat
                    end if
                else if mp3filename contains "60 radio" then
                repeat with doc1 in docxHar
                    if doc1 contains "60" then
                        if doc1 contains "radio" then
                        set docpath to doc1
                        exit repeat
                        end if
                    end if
                end repeat
                else if mp3filename contains "30 radio" then
                repeat with doc1 in docxHar
                    if doc1 contains "30" then
                        if doc1 contains "radio" then
                        set docpath to doc1
                        exit repeat
                        else if doc1 contains "digital" then
                        set docpath to doc1
                        exit repeat
                        end if
                    end if
                end repeat
                else if mp3filename contains "15 TV" then
                repeat with doc1 in docxHar
                    if doc1 contains "15" then
                        if doc1 contains "TV" then
                        set docpath to doc1
                        exit repeat
                        end if
                    end if
                end repeat
                else if mp3filename contains "30 TV" then
                repeat with doc1 in docxHar
                    if doc1 contains "30" then
                        if doc1 contains "TV" then
                        set docpath to doc1
                        exit repeat
                        end if
                    end if
                end repeat
                else if mp3filename contains "10 TV" then
                repeat with doc1 in docxHar
                    if doc1 contains "10" then
                        if doc1 contains "TV" then
                        set docpath to doc1
                        exit repeat
                        end if
                    end if
                end repeat
                else
                set searchtermsz to {}
                log "tried to open window"
                my openctwin1()
                set my theconz to true
                repeat while (my theconz) is true
                    delay 3
                end repeat
                my closectwin1()
                if (my r60cb1's state) as string is "1" then
                    set end of searchtermsz to "60 radio"
                end if
                if (my r30cb1's state) as string is "1" then
                    set end of searchtermsz to "30 radio"
                end if
                if (my t30cb1's state) as string is "1" then
                    set end of searchtermsz to "30 TV"
                end if
                if (my t15cb1's state) as string is "1" then
                    set end of searchtermsz to "15 TV"
                end if
                if (my t10cb1's state) as string is "1" then
                    set end of searchtermsz to "10 TV"
                end if
                if (my t60cb1's state) as string is "1" then
                    set end of searchtermsz to "60 TV"
                end if
                if (my rs60cb1's state) as string is "1" then
                    set end of searchtermsz to "60 radio Spanish"
                end if
                if (my rs30cb1's state) as string is "1" then
                    set end of searchtermsz to "30 radio Spanish"
                end if
                if (my ts30cb1's state) as string is "1" then
                    set end of searchtermsz to "30 TV Spanish"
                end if
                if (my ts15cb1's state) as string is "1" then
                    set end of searchtermsz to "15 TV Spanish"
                end if
                if (my ts10cb1's state) as string is "1" then
                    set end of searchtermsz to "10 TV Spanish"
                end if
                if (my ts60cb1's state) as string is "1" then
                    set end of searchtermsz to "60 TV Spanish"
                end if
                set (my r60cb1's state) to 0
                set (my r30cb1's state) to 0
                set (my t30cb1's state) to 0
                set (my t15cb1's state) to 0
                set (my t10cb1's state) to 0
                set (my t60cb1's state) to 0
                set (my rs60cb1's state) to 0
                set (my rs30cb1's state) to 0
                set (my ts30cb1's state) to 0
                set (my ts15cb1's state) to 0
                set (my ts10cb1's state) to 0
                set (my ts60cb1's state) to 0
                set docpath to {}
                repeat with type1 in searchtermsz
                    set type1ar to my theSplit(type1, " ")
                    set type11 to item 1 of type1ar
                    set type12 to item 2 of type1ar
                if type1 contains "Spanish" then
                repeat with doc1 in docxHar
                    if doc1 contains type11 then
                        if doc1 contains type12 then
                            if doc1 contains "AB" then
                        set end of docpath to doc1
                        exit repeat
                        end if
                        end if
                    end if
                end repeat
                else
                repeat with doc1 in docxHar
                    if doc1 contains type11 then
                        if doc1 contains type12 then
                        set end of docpath to doc1
                        exit repeat
                        end if
                    end if
                end repeat
                end if
                end repeat
                end if
                end if
            end try
            tell application "Microsoft Excel"
                tell workbook "Database.xlsx"
                    tell sheet "Main Base"
            set ranyggg to "E" & datalinef
            set clicon to value of range ranyggg as string
            my setsessmoddate(datalinef)
        if mp3filename does not contain "COMP VO" then
        if mp3filename contains " REV" then
            set buttonr to my roollrev(mp3filename)
            if buttonr is "Yes" then
                if mp3filename contains "TV"
                    set ranyggg to "D" & datalinef
                    set randydo to "F" & datalinef
                    if mp3filename contains "30" then
                    set subbytwo to (value of range ranyggg as string) & " :30 TV"
                    else if mp3filename contains "15" then
                    set subbytwo to (value of range ranyggg as string) & " :15 TV"
                    else if mp3filename contains "10" then
                    set subbytwo to (value of range ranyggg as string) & " :10 TV"
                    else
                    set subbytwo to (value of range ranyggg as string) & " :TV"
                    end if
                    set ranyggg to "I" & datalinef
                    set value of cell ranyggg to "Approved"
                    set subbytwo to subbytwo & " REV"
                    my sendoutREVTVa(theaiffpath, subbytwo, clicon, datalinef, replycheck, mp3filename)
                else
                set ranyggg to "D" & datalinef
                set randydo to "F" & datalinef
                if mp3filename contains "30" then
                set subbytwo to (value of range ranyggg as string) & " :30 Radio"
                else if mp3filename contains "60" then
                set subbytwo to (value of range ranyggg as string) & " :60 Radio"
                else
                set subbytwo to (value of range ranyggg as string) & " :Radio"
                end if
                set ranyggg to "I" & datalinef
                set value of cell ranyggg to "Approved"
                set subbytwo to subbytwo & " REV"
                my sendoutREVa(themp2path, subbytwo, clicon, datalinef, replycheck, mp3filename)
                end if
            else
            if mp3filename contains "TV"
                set ranyggg to "D" & datalinef
                set randydo to "F" & datalinef
                if mp3filename contains "30" then
                set subbytwo to (value of range ranyggg as string) & " :30 TV"
                else if mp3filename contains "15" then
                set subbytwo to (value of range ranyggg as string) & " :15 TV"
                else if mp3filename contains "10" then
                set subbytwo to (value of range ranyggg as string) & " :10 TV"
                else
                set subbytwo to (value of range ranyggg as string) & " :TV"
                end if
                set ranyggg to "I" & datalinef
                set value of cell ranyggg to "TV Mix sent"
                my sentouttvfora(theaiffpath, subbytwo, clicon, datalinef, replycheck, mp3filename)
            else
                set ranyggg to "D" & datalinef
                set randydo to "F" & datalinef
                if mp3filename contains "30" then
                set subbytwo to (value of range ranyggg as string) & " :30 Radio"
                else if mp3filename contains "60" then
                set subbytwo to (value of range ranyggg as string) & " :60 Radio"
                else
                set subbytwo to (value of range ranyggg as string) & " :Radio"
                end if
                set ranyggg to "I" & datalinef
                set value of cell ranyggg to "Radio mix sent"
                my sendoutmixfora(themp2path, subbytwo, clicon, datalinef, replycheck, mp3filename)
                end if
            end if
        else if mp3filename contains "radio" as string then
            set ranyggg to "D" & datalinef
            set randydo to "F" & datalinef
            if mp3filename contains "30" then
            set subbytwo to (value of range ranyggg as string) & " 30 Radio"
            else if mp3filename contains "60" then
            set subbytwo to (value of range ranyggg as string) & " 60 Radio"
            else
            set subbytwo to (value of range ranyggg as string) & " Radio"
            end if
            set ranyggg to "I" & datalinef
            set value of cell ranyggg to "Radio mix sent"
            my sendoutmixfora(themp2path, subbytwo, clicon, datalinef, replycheck, mp3filename)
        else if mp3filename contains "TV" as string then
            set ranyggg to "D" & datalinef
            set randydo to "F" & datalinef
            if mp3filename contains "30" then
            set subbytwo to (value of range ranyggg as string) & " 30 TV"
            else if mp3filename contains "15" then
            set subbytwo to (value of range ranyggg as string) & " 15 TV"
            else if mp3filename contains "10" then
            set subbytwo to (value of range ranyggg as string) & " 10 TV"
            else
            set subbytwo to (value of range ranyggg as string) & " TV"
            end if
            set ranyggg to "I" & datalinef
            set value of cell ranyggg to "TV Mix sent"
            log "bouncedfile function sentouttvfora"
            my sentouttvfora(themp2path, subbytwo, clicon, datalinef, replycheck, mp3filename)
        end if
        else if mp3filename contains "COMP VO" as string then
            set ranyggg to "D" & datalinef
            set randydo to "F" & datalinef
            set subbytwo to (value of range ranyggg as string)
            set ranyggg to "I" & datalinef
            set value of cell ranyggg to "COMP VO sent"
            my sendoutcomp(themp2path, docpath, subbytwo, clicon, datalinef, replycheck, mp3filename)
        end if
        end tell
        end tell
        end tell
    end bouncedfile
    on roollrev(mp3filename)
        display dialog "Roll This REV? " & mp3filename buttons {"Yes", "No"}
        if the button returned of the result is "Yes" then
            return("Yes")
        else
            return("No")
        end if
    end roollrev
    on CTbutpress_(Sender)
        set (my theconz) to false
    end CTbutpress_
    on replayed(datalinee)
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
                set rangzzz to (cell ("M") & datalinee)
                set value of rangzzz to ""
                end tell
                end tell
            end tell
    end replayed
    on sendoutcomp(themp2path, docpath, subbytwo, clicon, datalinee, replycheck, mp3filename)
        log "func: sendoutcomp"
        log "func: sendoutcomp var replycheck:" & replycheck as string
        if replycheck contains "COMP" as string then
            set replycheck to item 2 of my theSplit(replycheck, "COMP VO:")
            set replycheck to item 1 of my theSplit(replycheck, "*")
            tell application "Microsoft Outlook"
                    set theMsg to message id replycheck
                    set newMsg to reply to theMsg reply to all true
                    activate
                    my addattachz(themp2path)
                    set subbyz to subbytwo
                    set bodyz to "How is this?"
                    my setthesubabo(subbyz, bodyz)
                    
                    if count of docpath is greater than 10 then
                        my addattachz(docpath)
                    else if count of docpath is greater than 1 then
                        repeat with dov1 in docpath
                            my addattachz(dov1)
                        end repeat
                    else
                        set dov1 to docpath as string
                        my addattachz(dov1)
                    end if
                    my replayed(datalinee)
            end tell
        else
        if clicon contains "Max" then
            tell application "Finder"
                 open ("/Users/jonathanstoff/Documents/email templates/Max COMPVO.emltpl" as Posix file)
                delay 1
            end tell
                my addattachz(themp2path)
                set subbyz to subbytwo & " COMP VO"
                my setthesubabo(subbyz, subbytwo)
                if count of docpath is greater than 10 then
                    my addattachz(docpath)
                else if count of docpath is greater than 1 then
                    repeat with dov1 in docpath
                        my addattachz(dov1)
                    end repeat
                else
                    set dov1 to docpath as string
                    my addattachz(dov1)
                end if
                my setthestat("Comp vo sent to Max", datalinee, "COMP VO")
        else
            if mp3filename contains "span" as string then
                tell application "Finder"
                     open ("/Users/jonathanstoff/Documents/email templates/MAX AND TOM COMPVOspn.emltpl" as Posix file)
                    delay 1
                end tell
                    my addattachz(themp2path)
                    set subbyz to subbytwo & "SPANISH COMP VO"
                    my setthesubabo(subbyz, subbytwo)
                    if count of docpath is greater than 10 then
                        my addattachz(docpath)
                    else if count of docpath is greater than 1 then
                        repeat with dov1 in docpath
                            my addattachz(dov1)
                        end repeat
                    else
                        set dov1 to docpath as string
                        my addattachz(dov1)
                    end if
                    my setthestat("COMP VO sent to Max & Tom", datalinee, "COMP VO")
            else
            tell application "Finder"
             open ("/Users/jonathanstoff/Documents/email templates/TOM COMPVO.emltpl" as Posix file)
                delay 1
            end tell
                my addattachz(themp2path)
                set subbyz to subbytwo & " COMP VO"
                my setthesubabo(subbyz, subbytwo)
                #display dialog docpath as string
                if count of docpath is greater than 10 then
                    my addattachz(docpath)
                else if count of docpath is greater than 1 then
                    repeat with dov1 in docpath
                        my addattachz(dov1)
                    end repeat
                else
                    set dov1 to docpath as string
                    my addattachz(dov1)
                end if
                my setthestat("Comp VO sent to Tom", datalinee, "COMP VO")
            end if
            end if
            
        end if
    end sendoutcomp
    on sentouttvfora(themp2path, subbytwo, clicon, datalinee, replycheck, mp3filename)
        log "sentouttvfora vars: " & themp2path & subbytwo & clicon & datalinee & replycheck as string
        try
        set typeC to item 2 of my theSplit(subbytwo, ":")
        delay 1
        end try
        #display dialog "1"
        if replycheck contains "TV" then
            if mp3filename contains "30" then
                set temp1tv to "TV 30:"
            else if mp3filename contains "10" then
                set temp1tv to "TV 10:"
            else if mp3filename contains "15" then
                set temp1tv to "TV 15:"
            else
                set temp1tv to "TV:"
            end if
            set replycheck to item 2 of my theSplit(replycheck, temp1tv)
            set replycheck to item 1 of my theSplit(replycheck, "*")
                tell application "Microsoft Outlook"
                        set theMsg to message id replycheck
                        set newMsg to reply to theMsg reply to all true
                        activate
                        my addattachz(themp2path)
                        set subbyz to subbytwo
                        set bodyz to "How is this mix?"
                        my setthesubabo(subbyz, bodyz)
                        my replayed(datalinee)
                end tell
        else
        if clicon contains "Jerry" then
            #display dialog "2"
               tell application "Finder"
                   open ("/Users/jonathanstoff/Documents/email templates/Jerry mix.emltpl" as Posix file)
                    delay 1
                end tell
                    my addattachz(themp2path)
                    set subbyz to subbytwo
                    my setthesub(subbyz)
                    my setthestat("Mix sent to Jerry", datalinee, typeC)
        else if clicon contains "Max" then
            tell application "Finder"
                open ("/Users/jonathanstoff/Documents/email templates/Max mix.emltpl" as Posix file)
                 delay 1
             end tell
                 my addattachz(themp2path)
                 set subbyz to subbytwo
                 my setthesub(subbyz)
                 my setthestat("Mix sent to Max", datalinee, typeC)
                 else if (themp2path as string) contains "kia_of_c" then
                     tell application "Finder"
                         open ("/Users/jonathanstoff/Documents/email templates/Jerry mix.emltpl" as Posix file)
                         set subbyz to name of (themp2path as alias)
                         set subbyz to my theSplit(subbyz, ".mp3") as string
                          delay 1
                      end tell
                          my addattachz(themp2path)
                          
                          my setthesub(subbyz)
                          my setthestat("Mix sent to Jerry", datalinee, typeC)
        end if
        end if
    end sentouttvfora
    on sendoutmixfora(themp2path, subbytwo, clicon, datalinee, replycheck, mp3filename)
        try
        set typeC to item 2 of my theSplit(subbytwo, ":")
        delay 1
        end try
        log "sendoutmixfora fun start"
        set temp1tv to "radio:"
        if replycheck contains "radio 60" then
            if mp3filename contains "60" then
                set temp1tv to "radio 60:"
            end if
        end if
        if replycheck contains "radio 30" then
        if mp3filename contains "30" then
            set temp1tv to "radio 30:"
        end if
        end if
        if replycheck contains "radio" then
            set replycheck to item 2 of my theSplit(replycheck, temp1tv)
            set replycheck to item 1 of my theSplit(replycheck, "*")
                tell application "Microsoft Outlook"
                        set theMsg to message id replycheck
                        set newMsg to reply to theMsg reply to all true
                        activate
                        my addattachz(themp2path)
                        set subbyz to subbytwo
                        set bodyz to "How is this mix?"
                        my setthesubabo(subbyz, bodyz)
                        my replayed(datalinee)
                end tell
        else
        if clicon contains "Jerry" then
                tell application "Finder"
                    open ("/Users/jonathanstoff/Documents/email templates/Jerry mix.emltpl" as Posix file)
                 end tell
                     my addattachz(themp2path)
                     set subbyz to subbytwo
                     my setthesub(subbyz)
                     my setthestat("Mix sent to Jerry", datalinee, typeC)
        else if clicon contains "Max" then
            tell application "Finder"
                    open ("/Users/jonathanstoff/Documents/email templates/Max mix.emltpl" as Posix file)
             end tell
                 my addattachz(themp2path)
                 set subbyz to subbytwo
                 my setthesub(subbyz)
                 my setthestat("Mix sent to Max", datalinee, typeC)
        end if
        end if
    end sendoutmixfora
    on sendoutREVa(themp2path, subbytwo, clicon, datalinee, replycheck, mp3filename)
        try
        set typeC to item 2 of my theSplit(subbytwo, ":")
        delay 1
        end try
        if clicon contains "Jerry" then
                tell application "Finder"
                    open ("/Users/jonathanstoff/Documents/email templates/Approved Radio Jerry.emltpl" as Posix file)
                 end tell
                     my addattachz(themp2path)
                     set subbyz to subbytwo
                     my setthesub(subbyz)
                     my setthestat("Approved", datalinee, typeC)
        else if clicon contains "Max" then
            tell application "Finder"
                    open ("/Users/jonathanstoff/Documents/email templates/Approved Radio Max.emltpl" as Posix file)
             end tell
                 my addattachz(themp2path)
                 set subbyz to subbytwo
                 my setthesub(subbyz)
                 my setthestat("Approved", datalinee, typeC)
        end if
    end sendoutREVa
    on sendoutREVTVa(themp2path, subbytwo, clicon, datalinee, replycheck, mp3filename)
        try
        set typeC to item 2 of my theSplit(subbytwo, ":")
        delay 1
        end try
                tell application "Finder"
                    open ("/Users/jonathanstoff/Documents/email templates/Approved TV audio.emltpl" as Posix file)
                 end tell
                     my addattachz(themp2path)
                     set subbyz to subbytwo
                     my setthesub(subbyz)
                     my setthestat("Approved", datalinee, typeC)
    end sendoutREVTVa
    on openctwin1()
        log "Openctwin1"
        current application's ctwin1's OrderFront_()
    end openctwin1
    on closectwin1()
        log "Closectwin1"
        current application's ctwin1's OrderOut_()
    end closectwin1
    on approved(intt, fmesgfiveij, sendera)
        log "func approved"
        set sessdat to my getsessionnameup(fmesgfiveij)
        log "func approved var sessdat:" & (sessdat as string)
        set sessnamee to item 1 of sessdat
        set datalinee to item 2 of sessdat
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                    tell sheet "Main Base"
                        set typezran to (cell ("G" & datalinee))
                        set typesG to value of typezran
                        set rangz to "D" & datalinee
                        set rangzx to "F" & datalinee
                        set tempcli to value of cell rangz
                        set temptit to value of cell rangzx
                        set findthesesst to temptit & tempcli
                        set findthesess to my theSplit(findthesesst, " ") as string
                    end tell
                end tell
            end tell
            set typesar to {""}
            try
                set typesG to my trimthis(typesG, true, "full")
            end try
            try
                set typesar to my theSplit(typesG, "*")
            end try
            
            if count of typesar is greater than 1
                set searchtermsz to {}
                my ctwin2's makeKeyAndOrderFront_(1)
                set my theconz to true
                repeat while my theconz is true
                    delay 3
                end repeat
                my ctwin2's makeKeyAndOrderFront_(0)
                if (my r60cb1's state) as string is "1" then
                    set end of searchtermsz to "60 radio"
                end if
                if (my r30cb1's state) as string is "1" then
                    set end of searchtermsz to "30 radio"
                end if
                if (my t30cb1's state) as string is "1" then
                    set end of searchtermsz to "30 TV"
                end if
                if (my t15cb1's state) as string is "1" then
                    set end of searchtermsz to "15 TV"
                end if
                if (my t10cb1's state) as string is "1" then
                    set end of searchtermsz to "10 TV"
                end if
                if (my t60cb1's state) as string is "1" then
                    set end of searchtermsz to "60 TV"
                end if
                set (my r30cb1's state) to 0
                set (my r30cb1's state) to 0
                set (my t30cb1's state) to 0
                set (my t15cb1's state) to 0
                set (my t10cb1's state) to 0
                set (my t60cb1's state) to 0
                tell application "Microsoft Excel"
                tell workbook "Audio Production Sheet.xlsx"
                    tell sheet "Audio"
                        set searchydoi to 1
                        set docpath to {}
                        repeat with type1 in searchtermsz
                            set type1ar to my theSplit(type1, " ")
                            set type11 to item 1 of type1ar
                            set type12 to item 1 of type1ar
                        repeat 100 times
                            set rangt to "B" & searchydoi
                            set rangc to "A" & searchydoi
                            set ranzi to (value of cell rangt) & (value of cell rangc)
                            set strz to my theSplit(ranzi, " ") as string
                            if strz contains findthesess then
                                set rangzzz to cell ("C" & searchydoi)
                                if value of rangzzz as string contains type11 then
                                    if value of rangzzz as string contains type12 then
                                        set datelz to searchydoi
                                        set rangz to "A" & datelz & ":I" & datelz
                                        set value of range rangz to ""
                                        exit repeat
                                    end if
                                end if
                            end if
                        set searchydoi to searchydoi + 1
                        end repeat
                        end repeat
                    end tell
                end tell
                    tell workbook "Database.xlsx"
                        tell sheet "Main Base"
                #update database
                    set ranz7 to "P" & datalinee
                    set idemail to value of cell ranz7
                    set ranyg to "I" & datalinee
                    set value of cell ranyg to "Competed" as string
                    set thedaytoday1ar to current date
                    set thedaytoday2 to (month of thedaytoday1ar & "/" & day of thedaytoday1ar) as string
                    set ranyg to "N" & datalinee
                    set value of cell ranyg to thedaytoday2 as string
                #end update
                set ranyg to "C" & datalinee
                set ranyggg to "D" & datalinee
                set pathb to value of range ranyg & "Bounced Files:" as alias
                set mp3filename to {}
                set themp2path to {}
                set pathbb to "\"" & POSIX path of pathb & "\""
                repeat with type1 in searchtermsz
                    if type contains "Radio" then
                set end of mp3filename to do shell script "cd " & pathbb & " && ls -t *" & (type1 as string) & "*.mp3 | sed -n 1p"
                set end of themp2path to (pathb & mp3filename) as string
                else
                set end of mp3filename to do shell script "cd " & pathbb & " && ls -t *" & (type1 as string) & "*.aif | sed -n 1p"
                set end of themp2path to (pathb & mp3filename) as string
                end if
                end repeat
                #tell application "Finder" to reveal themp2path
                set ranygg to "E" & datalinee
                set clicon to value of range ranygg as string
                if mp3filename contains "radio" as string then
                    set subbytwo to value of range ranyggg & " radio"
                    if sendera contains "Jerry" then
                        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Approved Radio Jerry.emltpl" as Posix file)
                    else if sendera contains "Max" then
                    tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Approved Radio Max.emltpl" as Posix file)
                    else if clicon contains "Jerry" then
                    tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Approved Radio Jerry.emltpl" as Posix file)
                    else if clicon contains "Max" then
                    tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Approved Radio Max.emltpl" as Posix file)
                    end if
                else if mp3filename contains "TV" as string then
                    set subbytwo to value of range ranyggg & " TV"
                    set pathbb to "\"" & POSIX path of pathb & "\""
                    set mp3filename to do shell script "cd " & pathbb & " && ls -t *.aif | sed -n 1p" as string
                    set themp2path to (pathb & mp3filename) as string
                    tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Approved TV audio.emltpl" as Posix file)
                end if
                end tell
                end tell
                end tell
            else
            tell application "Microsoft Excel"
            tell workbook "Audio Production Sheet.xlsx"
                tell sheet "Audio"
                    set searchydoi to 1
                    set datelz to false
                    repeat 100 times
                        set rangt to "B" & searchydoi
                        set rangc to "A" & searchydoi
                        set ranzi to (value of cell rangt) & (value of cell rangc)
                        set strz to my theSplit(ranzi, " ") as string
                        if strz contains findthesess then
                            set datelz to searchydoi
                            exit repeat
                        end if
                    set searchydoi to searchydoi + 1
                    end repeat
                    if datelz is false then
                    else
                    set rangz to "A" & datelz & ":I" & datelz
                    set value of range rangz to ""
                    end if
                end tell
            end tell
                tell workbook "Database.xlsx"
                    tell sheet "Main Base"
            #update database
                set ranz7 to "P" & datalinee
                set idemail to value of cell ranz7
                set ranyg to "I" & datalinee
                set value of cell ranyg to "Competed" as string
                set thedaytoday1ar to current date
                set thedaytoday2 to (month of thedaytoday1ar & "/" & day of thedaytoday1ar) as string
                set ranyg to "N" & datalinee
                set value of cell ranyg to thedaytoday2 as string
            #end update
            set ranyg to "C" & datalinee
            set ranyggg to "D" & datalinee
            set pathb to value of range ranyg & "Bounced Files:" as alias
            
            set pathbb to "\"" & POSIX path of pathb & "\""
            set mp3filename to do shell script "cd " & pathbb & " && ls -t *.mp3 | sed -n 1p" as string
            set themp2path to (pathb & mp3filename) as string
            #tell application "Finder" to reveal themp2path
            set ranygg to "E" & datalinee
            set clicon to value of range ranygg as string
            if mp3filename contains "radio" as string then
                set subbytwo to value of range ranyggg & " radio"
                if sendera contains "Jerry" then
                    tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Approved Radio Jerry.emltpl" as Posix file)
                else if sendera contains "Max" then
                tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Approved Radio Max.emltpl" as Posix file)
                else if clicon contains "Jerry" then
                tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Approved Radio Jerry.emltpl" as Posix file)
                else if clicon contains "Max" then
                tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Approved Radio Max.emltpl" as Posix file)
                end if
            else if mp3filename contains "TV" as string then
                set subbytwo to value of range ranyggg & " TV"
                set pathbb to "\"" & POSIX path of pathb & "\""
                set mp3filename to do shell script "cd " & pathbb & " && ls -t *.aif | sed -n 1p" as string
                set themp2path to (pathb & mp3filename) as string
                tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Approved TV audio.emltpl" as Posix file)
            end if
            end tell
            end tell
                    end tell
        end if
        
                    try
                    set idsar to my theSplit(idemail, "*")
                    repeat with idstr in idsar
                    tell application "Microsoft Outlook" to move message id idstr of inbox of exchange account "spots@jerrydefalco.com" to mail folder "Archive" of exchange account "spots@jerrydefalco.com"
                        end repeat
                    end try
                my addattachz(themp2path)
                set subbyz to subbytwo
                my setthesub(subbyz)
    end approved
    on approved2(sessname, datalinee, tempcli, temptit, sendera)
                        log "func approved2"
                        tell application "Microsoft Excel"
                         set findthesesst to temptit & tempcli
                          set findthesess to my theSplit(findthesesst, " ") as string
                            tell workbook "Audio Production Sheet.xlsx"
                                tell sheet "Audio"
                                    set searchydoi to 1
                                    set datelz to false
                                    repeat 100 times
                                        set rangt to "B" & searchydoi
                                        set rangc to "A" & searchydoi
                                        set ranzi to (value of cell rangt) & (value of cell rangc)
                                        set strz to my theSplit(ranzi, " ") as string
                                        if strz contains findthesess then
                                            set datelz to searchydoi
                                            exit repeat
                                        end if
                                    set searchydoi to searchydoi + 1
                                    end repeat
                                    if datelz is false then
                                    else
                                    set rangz to "A" & datelz & ":I" & datelz
                                    set value of range rangz to ""
                                    end if
                                end tell
                            end tell
                                tell workbook "Database.xlsx"
                                    tell sheet "Main Base"
                            #update database
                                set ranz7 to "P" & datalinee
                                set idemail to value of cell ranz7
                                set ranyg to "I" & datalinee
                                set value of cell ranyg to "Competed" as string
                                set thedaytoday1ar to current date
                                set thedaytoday2 to (month of thedaytoday1ar & "/" & day of thedaytoday1ar) as string
                                set ranyg to "N" & datalinee
                                set value of cell ranyg to thedaytoday2 as string
                            #end update
                            set ranyg to "C" & datalinee
                            set ranyggg to "D" & datalinee
                            set pathb to value of range ranyg & "Bounced Files:" as alias
                            set pathbb to "\"" & POSIX path of pathb & "\""
                            set mp3filename to do shell script "cd " & pathbb & " && ls -t *.mp3 | sed -n 1p" as string
                            set themp2path to (pathb & mp3filename) as string
                            #tell application "Finder" to reveal themp2path
                            set ranygg to "E" & datalinee
                            set clicon to value of range ranygg as string
                            if mp3filename contains "radio" as string then
                                set subbytwo to value of range ranyggg & " radio"
                                if sendera contains "Jerry" then
                                    tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Approved Radio Jerry.emltpl" as Posix file)
                                else if sendera contains "Max" then
                                tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Approved Radio Max.emltpl" as Posix file)
                                else if clicon contains "Jerry" then
                                tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Approved Radio Jerry.emltpl" as Posix file)
                                else if clicon contains "Max" then
                                tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Approved Radio Max.emltpl" as Posix file)
                                end if
                            else if mp3filename contains "TV" as string then
                                set subbytwo to value of range ranyggg & " TV"
                                set pathbb to "\"" & POSIX path of pathb & "\""
                                set mp3filename to do shell script "cd " & pathbb & " && ls -t *.aif | sed -n 1p" as string
                                set themp2path to (pathb & mp3filename) as string
                                tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Approved TV audio.emltpl" as Posix file)
                            end if
                            end tell
                            end tell
                        end tell
                                    try
                                    set idsar to my theSplit(idemail, "*")
                                    repeat with idstr in idsar
                                    tell application "Microsoft Outlook" to move message id idstr of inbox of exchange account "spots@jerrydefalco.com" to mail folder "Archive" of exchange account "spots@jerrydefalco.com"
                                        end repeat
                                    end try
                                my addattachz(themp2path)
                                set subbyz to subbytwo
                                my setthesub(subbyz)
                    end approved2
    on Compvogood(msg)
    log "Compvogood"
    tell application "Microsoft Outlook"
    set newestmm to msg
    set subbyjj to subject of newestmm as string
    set subbyjjar to my theSplit(subbyjj, "Re: ")
    set subbyjj to subbyjjar as string
    set bodyz to plain text content of newestmm
    end tell
    try
    set bodyzar to my theSplit(bodyz, "Tom Rowe")
    set newbod to item 1 of bodyzar
    end try
    set sessdat to my getsessionnameup(msg)
    set sessnamee to item 1 of sessdat
    set datalinee to item 2 of sessdat
    tell application "Microsoft Excel"
        tell workbook "Database.xlsx"
            tell sheet "Main Base"
                set ranzx to "C" & datalinee
                set fold1 to (value of cell ranzx as string) & "Bounced Files:"
                set ranzx to "H" & datalinee
                set docxH to (value of cell ranzx as string)
            end tell
        end tell
    end tell
    try
        set docxHar to my theSplit(docxH, "*")
        repeat with doc1 in docxHar
            if doc1 contains "60" then
                set docxH to doc1
                exit repeat
            else if doc1 contains "TV"
                if doc1 contains "30"
                    set docxH to doc1
                    exit repeat
                end if
            end if
        end repeat
    end try
    set pathbb to "\"" & POSIX path of fold1 & "\""
    set mp3filename to do shell script "cd " & pathbb & " && ls -t *.mp3 | sed -n 1p" as string
    set themp2path to fold1 & mp3filename
    tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/JERRY COMPVO.emltpl" as Posix file)
    my addattachz(themp2path)
    my addattachz(docxH)
    set bodz to subbyjj
    set subbycjj to subbyjj
    my setthesubabo(subbycjj, bodz)
    my setthestat("COMP VO sent to Jerry", datalinee, "COMP VO")
    log "fun compvogood " & datalinee
    end Compvogood
    
    on Compvosugcut(msg)
    tell application "Microsoft Outlook"
        set newestmm to msg
        set subbyjj to subject of newestmm
        set subbyjjar to my theSplit(subbyjj, "Re: ")
        set subbyjj to subbyjjar as string
        set bodyz to plain text content of newestmm as string
    end tell
        try
        set bodyzar to my theSplit(bodyz, "Tom Rowe")
        set newbod to item 1 of bodyzar
        end try
        set sessdat to my getsessionnameup(msg)
        set sessnamee to item 1 of sessdat
        set datalinee to item 2 of sessdat
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
                    set ranzx to "C" & datalinee
                    set fold1 to (value of cell ranzx as string) & "Bounced Files:"
                    set ranzx to "H" & datalinee
                    set docxH to (value of cell ranzx as string)
                end tell
            end tell
        end tell
        try
            set docxHar to my theSplit(docxH, "*")
            repeat with doc1 in docxHar
                if doc1 contains "60" then
                    set docxH to doc1
                    exit repeat
                else if doc1 contains "TV"
                    if doc1 contains "30"
                        set docxH to doc1
                        exit repeat
                    end if
                end if
            end repeat
        end try
        set pathbb to "\"" & POSIX path of fold1 & "\""
        set mp3filename to do shell script "cd " & pathbb & " && ls -t *.mp3 | sed -n 1p" as string
        set themp2path to fold1 & mp3filename
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/JERRY COMPVO.emltpl" as Posix file)
        my addattachz(themp2path)
        my addattachz(docxH)
        set bodz to subbyjj & "                                                     Tom's suggested cuts:                                           " & newbod
        set subbycjj to subbyjj
        my setthesubabo(subbycjj, bodz)
        my setthestat("COMP VO sent to Jerry", datalinee, "COMP VO")
    end Compvosugcut
    on makecuts(theMsg)
    tell application "Microsoft Outlook"
        set newestmm to theMsg
        set idz to id of theMsg
        set subbyjj to subject of newestmm as string
        set bodyz to plain text content of newestmm
        set subbyx to subbyjj
        if subbyx contains "radio"
            if subbyx contains "60"
                set idz to "*radio 60:" & idz
            else if subbyx contains "30"
                set idz to "*radio 30:" & idz
            else
                set idz to "*radio:" & idz
            end if
        else if subbyx contains "TV"
            if subbyx contains "30"
                set idz to "*TV 30:" & idz
            else if subbyx contains "15"
                set idz to "*TV 15:" & idz
            else if subbyx contains "10"
                set idz to "*TV 10:" & idz
            else
                set idz to "*TV:" & idz
            end if
        else if subbyx contains "COMP VO"
            set idz to "*COMP VO:" & idz
        end if
    end tell
    with timeout of 86400 seconds
        set boolent to true
        repeat while boolent is true
            display dialog "Cuts for: " & subbyjj buttons {"Open session", "Wait 10 mins", "Got it!"}
            if the button returned of the result is "Open session" then
                set boolent to false
                exit repeat
            else if the button returned of the result is "Wait 10 mins" then
                delay 600
            else
                set boolent to true
                exit repeat
            end if
        end repeat
    end timeout
    set sessdat to my getsessionnameup(newestmm)
    set sessnamee to item 1 of sessdat
    set datalinee to item 2 of sessdat
    set sessionnamear to my getopensessname()
    set opensessnamee to item 1 of sessionnamear as string
    set nosess to item 2 of sessionnamear
    tell application "Microsoft Excel"
        tell workbook "Database.xlsx"
            tell sheet "Main Base"
    set ranyg to "M" & datalinee
    set value of cell ranyg to idz as string
    set thedaytoday1ar to current date
    set thedaytoday2 to (month of thedaytoday1ar & "/" & day of thedaytoday1ar) as string
    set ranyg to "N" & datalinee
    set value of cell ranyg to thedaytoday2 as string
    end tell
    end tell
    end tell
    if boolent is false
    if opensessnamee contains sessnamee then
        tell application "Finder"
            set subby to subbyjj & "_Edits" & ".txt"
            
                set filenameeq to "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" & subby
                delete filenameeq
                make new file at "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" with properties {name:subby, file type:"TEXT", creator type:"ttxt"}
                set tFile to open for access filenameeq with write permission
                try
                   write (stremail as string) to tFile starting at eof
                   close access tFile
                on error
                   close access tFile
                end try
                open filenameeq
            set filenameeq to "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" & subby
            open filenameeq
            end tell
    else
        opensessfromline(datalinee)
        delay 5
        tell application "Finder"
            set subby to subbyjj & "_Edits" & ".txt"
            
                set filenameeq to "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" & subby
                delete filenameeq
                make new file at "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" with properties {name:subby, file type:"TEXT", creator type:"ttxt"}
                set tFile to open for access filenameeq with write permission
                try
                   write (stremail as string) to tFile starting at eof
                   close access tFile
                on error
                   close access tFile
                end try
                open filenameeq
            set filenameeq to "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" & subby
            open filenameeq
            end tell
    end if
    else
    tell application "Finder"
        set subby to subbyjj & "_Edits" & ".txt"
        
            set filenameeq to "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" & subby
            delete filenameeq
            make new file at "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" with properties {name:subby, file type:"TEXT", creator type:"ttxt"}
            set tFile to open for access filenameeq with write permission
            try
               write (stremail as string) to tFile starting at eof
               close access tFile
            on error
               close access tFile
            end try
            open filenameeq
        set filenameeq to "Macintosh HD:var:folders:zh:q1x9hdq15ys9nx6c5shwp6pw0000gn:T:" & subby
        open filenameeq
        end tell
    end if
    end makecuts
    on change_case(this_text)
        set new_text to do shell script "echo " & quoted form of (this_text) & " | tr a-z A-Z"
        return new_text
    end change_case
    on change_cased(this_text)
        set new_text to do shell script "echo " & quoted form of (this_text) & " | tr A-Z a-z"
        return new_text
    end change_cased
    on readFile(unixPath)
    tell application "Finder"
        set foo to (open for access unixPath)
        #display dialog "1"
        set txt to (read foo for (get eof foo))
        #display dialog "2"
        close access foo
        
        return txt
        end tell
    end readFile
    on theSplit(theString, theDelimiter)
        #display dialog theString
        set oldDelimiters to AppleScript's text item delimiters
        set AppleScript's text item delimiters to theDelimiter
        set theArray to every text item of theString
        set AppleScript's text item delimiters to oldDelimiters
        return theArray
    end theSplit
    on trimThis(pstrSourceText, pstrCharToTrim, pstrTrimDirection)
        set strTrimedText to pstrSourceText
        if pstrCharToTrim is true then
            set pstrCharToTrim to {" ", tab, ASCII character 10, return, ASCII character 0}
        end if
        
        --- TRIM LEFT SIDE OF STRING ---
        
        if (pstrTrimDirection = "full") or (pstrTrimDirection = "left") then
            set iLoc to 1
            repeat until character iLoc of strTrimedText is not in pstrCharToTrim
                set iLoc to iLoc + 1
            end repeat
            
            set strTrimedText to text iLoc thru -1 of strTrimedText
        end if
        
        --- TRIM RIGHT SIDE OF STRING ---
        
        
        if (pstrTrimDirection = "full") or (pstrTrimDirection = "right") then
            set iLoc to count of strTrimedText
            repeat until character iLoc of strTrimedText is not in pstrCharToTrim
                set iLoc to iLoc - 1
            end repeat
            
            set strTrimedText to text 1 thru iLoc of strTrimedText
            
        end if
        
        return strTrimedText
        
    end trimThis
    on getsessionname(emailm)
        tell application "Microsoft Outlook"
            set subjem to subject of emailm as string
            try
                set subjem to my theSplit(subjem, "Re: ") as string
                set subjem to my theSplit(subjem, " COMP VO") as string
                set subjem to item 1 of my theSplit(subjem, "radio")
                set subjem to my theSplit(subjem, "60 ") as string
                set subjem to my theSplit(subjem, "TV") as string
                set subjem to item 1 of my theSplit(subjem, "30 ")
                set subjem to item 1 of my theSplit(subjem, " Mix")
                set subjem to item 1 of my theSplit(subjem, "15 ")
                set subjem to my trimThis(subjem, true, "full")
            end try
        end tell
        tell application "Microsoft Excel"
                tell workbook "Audio Production Sheet.xlsx"
                    tell sheet "Audio"

            set innt to 3
            set boolent to true
            repeat while boolent is true
                set myrgo to "A" & innt
                set myrgstr to value of range myrgo as string
                if myrgstr contains (subjem as string) then
                    set theinty to innt
                    set boolent to false
                else if innt is greater than 50 then
                    display dialog "failed to find in production sheet"
                end if
                set innt to innt + 1
            end repeat
            set myrangfo to "A" & theinty
            set myrangft to "B" & theinty
            set sessnlook to (value of range myrangft & " " & value of range myrangfo) as string
            try
            set sessnlook to my theSplit(sessnlook, "(Max)") as string
            set sessnlook to my theSplit(sessnlook, "(Jerry)") as string
            set sessnlook to my trimThis(sessnlook, true, "full")
            end try
            end tell
            end tell
                tell workbook "Database.xlsx"
                    tell sheet "Main Base"
            set innt to 2
            set boolent to true
            repeat while boolent is true
                set myrgo to "A" & innt
                if (value of range myrgo) as string contains (sessnlook as string) then
                    set dataline to innt
                    set boolent to false
                end if
                
                set innt to innt + 1
                
            end repeat
            set myrangf to "A" & dataline
            set sessnme to value of range myrangf as string
            set sessnme to my theSplit(sessnme, ".ptx") as string
            set outputz to {sessnme, dataline}
            end tell
            end tell
        end tell
        return (outputz)
    end getsessionname
    on getsessionnameup(emailm)
        log "Func: getsessionnameup"
        #log emailm
        set menuitemz to (my dynamicMenu)
        set listz to menuitemz's itemTitles() as list
        tell application "Microsoft Outlook"
            set subjem to subject of emailm
            set subjem to subjem as string
            try
                set subjem to my theSplit(subjem, "Re: ") as string
                set subjem to my theSplit(subjem, " COMP VO") as string
                set subjem to item 1 of my theSplit(subjem, "radio")
                set subjem to my theSplit(subjem, "60 ") as string
                set subjem to my theSplit(subjem, "TV") as string
                set subjem to item 1 of my theSplit(subjem, "30 ")
                set subjem to item 1 of my theSplit(subjem, " Mix")
                set subjem to item 1 of my theSplit(subjem, "15 ")
                set subjem to item 1 of my theSplit(subjem, "-")
                set subjem to my trimThis(subjem, true, "_")
                set subjem to my trimThis(subjem, true, "full")
            end try
        end tell
        log subjem as string
        try
        repeat with mitem in listz
            set mitem to my theSplit(mitem, " |") as string
            if mitem contains subjem then
                set themitem to mitem
                exit repeat
            end if
        end repeat
        set sessnme to item 1 of my theSplit(themitem, " |")
        set dataliz to item 2 of my theSplit(themitem, "<")
        set dataline to item 1 of my theSplit(dataliz, ">")
        set outputz to {sessnme, dataline}
        on error
        set outputz to {0, 0}
        end try
        log "getsessionnameup var output: " & outputz as string
        return (outputz)
    end getsessionnameup
    on getsessionnameupfz(emailm)
        set thisthing to true
        set theArray1 to {}
        repeat while thisthing is true
            log "Func: getsessionnameupfz"
            #log emailm
            set menuitemz to (my dynamicMenu)
            set listz to menuitemz's itemTitles() as list
        tell application "Microsoft Outlook"
            set subjem to subject of emailm
            set subjem to subjem as string
            try
                set subjem to my theSplit(subjem, "Re: ") as string
                set subjem to my theSplit(subjem, " COMP VO") as string
                set subjem to item 1 of my theSplit(subjem, "radio")
                set subjem to my theSplit(subjem, "60 ") as string
                set subjem to my theSplit(subjem, "TV") as string
                set subjem to item 1 of my theSplit(subjem, "30 ")
                set subjem to item 1 of my theSplit(subjem, " Mix")
                set subjem to item 1 of my theSplit(subjem, "15 ")
                set subjem to item 1 of my theSplit(subjem, "-")
                set subjem to my trimThis(subjem, true, "_")
                set subjem to my trimThis(subjem, true, "full")
            end try
        end tell
        log subjem as string
        try
            set theArray12 to theArray1 as string
        repeat with mitem in listz
            if theArray12 contains mitem then
            else
            if mitem contains subjem then
                set themitem to mitem
                exit repeat
            end if
            end if
        end repeat
        set sessnme to item 1 of my theSplit(themitem, " |")
        set dataliz to item 2 of my theSplit(themitem, "<")
        set dataline to item 1 of my theSplit(dataliz, ">")
        set outputz to {sessnme, dataline}
        on error
        set outputz to {0, 0}
        log "Error in getsessionnameupfz"
        end try
        set notfoundy to false
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
                    set theRanz to (cell ("G" & dataline))
                    set theType to (value of theRanz) as string
                    if theType does not contain "30 TV"
                        if theType does not contain "60"
                        set end of theArray1 to dataline
                        set notfoundy to true
                        end if
                    end if
                    if notfoundy is false
                    set thisthing to false
                    exit repeat
                    end if
                end tell
            end tell
        end tell
        end repeat
        log "Func: getsessionnameup var outputz: " & outputz as string
        return (outputz)
    end getsessionnameupfz
    on getsessionnameupz(emailm, subbyx)
        log "Func: getsessionnameup"
        log subbyx
        set menuitemz to (my dynamicMenu)
        set listz to menuitemz's itemTitles() as list
        tell application "Microsoft Outlook"
            set subjem to subbyx
            try
                set subjem to my theSplit(subjem, "Re: ") as string
                set subjem to my theSplit(subjem, " COMP VO") as string
                set subjem to item 1 of my theSplit(subjem, "radio")
                set subjem to my theSplit(subjem, "60 ") as string
                set subjem to my theSplit(subjem, "TV") as string
                set subjem to item 1 of my theSplit(subjem, "30 ")
                set subjem to item 1 of my theSplit(subjem, " Mix")
                set subjem to item 1 of my theSplit(subjem, "15 ")
                set subjem to item 1 of my theSplit(subjem, "-")
                set subjem to my trimThis(subjem, true, "_")
                set subjem to my trimThis(subjem, true, "full")
            end try
        end tell
        log subjem
        try
        repeat with mitem in listz
            set mitem1 to my theSplit(mitem, " |") as string
            if mitem1 contains subjem then
                set themitem to mitem
                exit repeat
            end if
        end repeat
        set sessnme to item 1 of my theSplit(themitem, " |")
        set dataliz to item 2 of my theSplit(themitem, "<")
        set dataline to item 1 of my theSplit(dataliz, ">")
        set outputz to {sessnme, dataline}
        on error
        set outputz to {0, 0}
        end try
        return (outputz)
    end getsessionnameupz
    on getsessionnameold(emailm)
        tell application "Microsoft Outlook"
            set subjem to subject of emailm as string
            try
                set subjem to my theSplit(subjem, "Re: ") as string
                set subjem to my theSplit(subjem, " COMP VO") as string
                set subjem to item 1 of my theSplit(subjem, "radio")
                set subjem to my theSplit(subjem, "60 ") as string
                set subjem to my theSplit(subjem, "TV") as string
                set subjem to item 1 of my theSplit(subjem, "30 ")
                set subjem to item 1 of my theSplit(subjem, " Mix")
                set subjem to item 1 of my theSplit(subjem, "15 ")
                set subjem to my trimThis(subjem, true, "full")
            end try
        end tell
        tell application "Microsoft Excel"
            set sessnlook to subjem
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
            set innt to 2
            set boolent to true
            repeat while boolent is true
                set myrgo to "A" & innt
                if (value of range myrgo) as string contains (sessnlook as string) then
                    set dataline to innt
                    set boolent to false
                end if
                
                set innt to innt + 1
                
            end repeat
            set myrangf to "A" & dataline
            set sessnme to value of range myrangf as string
            set sessnme to my theSplit(sessnme, ".ptx") as string
            set outputz to {sessnme, dataline}
            end tell
            end tell
        end tell
        return (outputz)
    end getsessionnameold
    on getopensessname()
        set nosess to false
        set sessionname to ""
        try
            tell application "System Events"
                tell process "Pro Tools"
                    set idmaker to (get name of every window)
                    #display dialog idmaker as string
                    set intz to count idmaker
                    set intint1 to 1
                    repeat intz times
                        set winz1 to item intint1 of idmaker
                        if winz1 contains "Edit" then
                            set sessionz to winz1
                            exit repeat
                        end if
                        set intint1 to intint1 + 1
                    end repeat
                end tell
                set sessionname to my theSplit(sessionz, "Edit: ") as string
            end tell
        on error
        set nosess to true
        end try
        return (sessionname, nosess)
    end getopensessname
    on opensessfromline(lined_)
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
            set ranyg to "B" & lined_
            set pathb to value of range ranyg as alias
            tell application "Finder" to open pathb
                end tell
            end tell
        end tell
        delay 1
        tell application "Keyboard Maestro Engine"
            do script "4912A6DF-5DCE-4DC7-A5CF-CC095EF409BD"
        end tell
    end opensessfromline
    #gets sess folder
    on opensessffromline(lined_)
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
            delay 2
            set ranyg to "C" & lined_
            
            set bleechs to value of range ranyg
            log bleechs as string
            return (bleechs as string)
        end tell
        end tell
        end tell
    end opensessffromline

    #Function to scan doc and output clientname, type, titlename, and announcers 1.clientname 2.clientcontact 3.titlename 4.spottype 5.announces
    on scanndoc(filedir)
    set failededed to false
        tell application "Finder"
            try
            move filedir to ("JDA_DATA:current scripts :" as alias)
            on error
                set failededed to true
            end try
            open filedir
        end tell
        if failededed is true then
            log "Failed to move doc to currentscripts"
        end if
        delay 3
        tell application "Microsoft Word"
            set theDoc to name of active document
            set theWin to name of front window
            set theSelectionn to selection
            tell selection
                delay 1
                set selection start to 2
                set selection end to 2000
                set document1 to content as string
                #display dialog content as string
                set lines1 to my theSplit(document1, "TITLE:")
                set line1 to item 1 of lines1
                set line2 to item 2 of lines1
                set lines2 to my theSplit(line2, "NEED")
                set line2 to item 1 of lines2
                set line3 to item 2 of lines2
                set lines3 to my theSplit(line3, "MUSIC:")
                set line3 to item 1 of lines3
                set line1array to my theSplit(line1, "spot:")
                #display dialog list2string(line1array, ",")
                set clientname2 to item 1 of line1array as text
                set spottype to item 2 of line1array as text
                set line1array to my theSplit(clientname2, "    ")
                set clientname to item 1 of line1array as text
                set line2array to my theSplit(line2, "NEED:")
                set titlename1 to item 1 of line2array as text
                set line2array to my theSplit(titlename1, "    ")
                set titlename to item 1 of line2array as text
                set titlename to my trimThis(titlename, true, "full")
                set clientname to my trimThis(clientname, true, "full")
                set spottype to my trimThis(spottype, true, "full")
                    
                set clientname1 to my theSplit(clientname, ": ")
                set clientname to item 2 of clientname1
                set clientname to my trimThis(clientname, true, "full")
            end tell
            
            
            if line3 contains "mark" then
                set end of announcervar to "Mark "
            end if
            if line3 contains "ben" then
                set end of announcervar to "Ben "
            end if
            if line3 contains "brent" then
                set end of announcervar to "Brent "
            end if
            if line3 contains "david" then
                set end of announcervar to "David "
            end if
            if line3 contains "doak" then
                set end of announcervar to "Doak "
            end if
            if line3 contains "melissa" then
                set end of announcervar to "Melissa "
            end if
            if line3 contains "mike" then
                set end of announcervar to "Mike o "
            end if
            if line3 contains "rachel" then
                set end of announcervar to "Rachel "
            end if
            if line3 contains "sandra" then
                set end of announcervar to "Sandra "
            end if
            if line3 contains "spanish" then
                set announcervar to "Spanish"
            end if
            if line2 contains "spn" then
                set announcervar to "Spanish"
            end if
            if line2 contains "spanish" then
                set announcervar to "Spanish"
            end if
            
        end tell
        set jerrysclientsl to my readFile("/Volumes/LaCie/Work/jclients.txt") as string
        if clientname contains "MATHEW" then
            set clientname to "MATHEWS AUTO GROUP"
        else if clientname contains "BENZ" then
            set clientname to "MERCEDES BENZ OF FT WAYNE"
        else if clientname contains "OPELIKA" then
            set clientname to "OPELIKA FORD CDJR"
        else if clientname contains "STRONG" then
            set clientname to "STRONG VW"
        else if clientname contains "PARK" then
            set clientname to "AUTO WORLD MITSUBISHI"
        else if clientname contains "CAPE" then
            set clientname to "CAPE CORAL KIA"
        end if
        
        if jerrysclientsl contains clientname then
            set clientcE to "Jerry's"
        else
            set clientcE to "Jerry's"
        end if
        set clientname to do shell script "echo " & quoted form of (clientname) & " | tr A-Z a-z"
        if failededed is true then
            try
            tell application "Finder" to move filedir to ("JDA_DATA:current scripts :" as alias) with replacing
            on error
            log "failed to move doc to JDA_DATA:current scripts :, you should try to do this yourself"
            end try
        end if
        return {clientname, clientcE, titlename, spottype, announcervar}
    end scanndoc
    #list1 is newest list, list2 is that last one
    on checkdiflist(list1, list2)
    tell application "Microsoft Outlook"
        set repeatbool to true
        set newmail to {}
        set typexz to count list1
        set typexzz to count list2
        if typexzz is 0 then
         set newmail to list1
        else
        set itemzz to 1
        set itemzzz to 1
        repeat typexz times
            set list1i to item itemzz of list1
            set foundsame to false
            repeat typexzz times
                set list2i to item itemzzz of list2
                if list1i is list2i then
                    set foundsame to true
                end if
                if itemzzz = typexzz then
                    if foundsame is false then
                        set end of newmail to list1i
                    end if
                end if
                set itemzzz to itemzzz + 1
            end repeat
            set itemzzz to 1
            set itemzz to itemzz + 1
        end repeat
        end if
        return (newmail)
        end tell
    end checkdiflist
    on searchdata(termy)
        set dataline to ""
        
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
                
            set innt to 2
            set boolent to true
            repeat while boolent is true
                
                set myrgo to "B" & innt
                set myrgog to "D" & innt
                if value of range myrgo contains (termy as string) then
                    set dataline to innt
                    set boolent to false
                    
                    exit repeat
                else if value of range myrgo is "" then
                if value of range myrgog is "" then
                    
                    set boolent to false
                    exit repeat
                    end if
                end if
                
                set innt to innt + 1
            end repeat
            
            return (dataline)
                end tell
            end tell
        end tell
        
    end searchdata
    on searchdataz(termy)
                        set dataline to ""
            
                        tell application "Microsoft Excel"
                            tell workbook "Database.xlsx"
                                tell sheet "Main Base"
                    
                            set innt to 2
                            set boolent to true
                            repeat while boolent is true
                                set myrgoz to cell ("A" & innt)
                                set myrgo to cell ("F" & innt)
                                set myrgog to cell ("D" & innt)
                                set linet to ((value of myrgo) & " " & (value of myrgog)) as string
                                set linet to my theSplit(linet, " ") as string
                                if linet contains (termy as string) then
                                    set dataline to innt
                                    set boolent to false
                                    exit repeat
                                else if linet is "" then
                                    if (value of myrgoz) as string is ""
                                        set dataline to innt
                                        set boolent to false
                                        exit repeat
                                    end if
                                end if
                                
                                set innt to innt + 1
                            end repeat
                            
                            return (dataline)
                                end tell
                            end tell
                        end tell
            
                    end searchdataz
    on setsessmoddate(datalinee)
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
        set thedaytoday1ar to current date
        set thedaytoday2 to (month of thedaytoday1ar & "/" & day of thedaytoday1ar) as string
        set ranyg to "N" & datalinee
        set value of cell ranyg to thedaytoday2 as string
        end tell
        end tell
        end tell
    end setsessmoddate
    on addtdbsh(sessnameA, sesspathB, sessfoldC, typesG, statusI, reciauL, sessmoddateN)
    tell application "Microsoft Excel"
        tell workbook "Database.xlsx"
            tell sheet "Main Base"
        set linerint to 2
        set newcellnum to 0
        set searchydo to false
        repeat while searchydo is false
            set linerange to "A" & linerint & ":F" & linerint
            set valuesz to value of range linerange as string
            
            if valuesz is "" then
                set newcellnum to linerint
                set searchydo to true
                exit repeat
            end if
            set linerint to linerint + 1
            
        end repeat
        set cellnumber to newcellnum
        set thedaytoday1ar to current date
        set thedaytoday2 to (month of thedaytoday1ar & "/" & day of thedaytoday1ar) as string
        set myrange to "A" & cellnumber
        set value of cell myrange to sessnameA as string
        set myrange to "B" & cellnumber
        set value of cell myrange to sesspathB as string
        set myrange to "C" & cellnumber
        set value of cell myrange to sessfoldC as string
        set myrange to "G" & cellnumber
        set value of cell myrange to typesG as string
        set myrange to "I" & cellnumber
        set value of cell myrange to statusI as string
        set myrange to "L" & cellnumber
        set value of cell myrange to reciauL as string
        set myrange to "N" & cellnumber
        set value of cell myrange to thedaytoday2 as string
        
        set value of cell myrange to thedaytoday2
    end tell
    end tell
    end tell
    end addtdbsh
            on addcurrdoc(sessnameA, sesspathB, sessfoldC, typesG, statusI, reciauL, sessmoddateN)
            tell application "Microsoft Excel"
                tell workbook "Database.xlsx"
                    tell sheet "Main Base"
                set linerint to 2
                set newcellnum to 0
                set searchydo to false
                repeat while searchydo is false
                    set linerange to "A" & linerint & ":F" & linerint
                    set valuesz to value of range linerange as string
                    
                    if valuesz is "" then
                        set newcellnum to linerint
                        set searchydo to true
                        exit repeat
                    end if
                    set linerint to linerint + 1
                    
                end repeat
                set cellnumber to newcellnum
                set thedaytoday1ar to current date
                set thedaytoday2 to (month of thedaytoday1ar & "/" & day of thedaytoday1ar) as string
                set myrange to "A" & cellnumber
                set value of cell myrange to sessnameA as string
                set myrange to "B" & cellnumber
                set value of cell myrange to sesspathB as string
                set myrange to "C" & cellnumber
                set value of cell myrange to sessfoldC as string
                set myrange to "G" & cellnumber
                set value of cell myrange to typesG as string
                set myrange to "I" & cellnumber
                set value of cell myrange to statusI as string
                set myrange to "L" & cellnumber
                set value of cell myrange to reciauL as string
                set myrange to "N" & cellnumber
                set value of cell myrange to thedaytoday2 as string
                
                set value of cell myrange to thedaytoday2
            end tell
            end tell
            end tell
            end addcurrdoc
            
    on writetopsess(additemnew)
        set menuitemz to (my dynamicMenu)
        set addeditemsz to additemnew
        menuitemz's addItemWithTitle_(addeditemsz)
        set listz to menuitemz's itemTitles() as list
    #menuitemz's release()
    tell application "Finder"
        set tempdirz to my psessdir
        set tempdirx to my zsessdir
        delete file tempdirz
        make new file at tempdirx with properties {name:"Previous_sessions.txt", file type:"TEXT", creator type:"ttxt"}
        set the open_target_file to (open for access (tempdirz) with write permission)
        
        set the_count to count listz
        write "" to open_target_file
        
        set in_t1 to 1
        repeat the_count times
            
            set str_z to (item in_t1 of listz) & "*"
            write str_z to open_target_file starting at eof
            set in_t1 to in_t1 + 1
        end repeat
        close access open_target_file
        end tell
    end writetopsess
            on formatDate(aDate)
                tell aDate to tell 100000000 + day * 1000000 + (its month) * 10000 + year as string ¬
                    to return text 4 thru 5 & "-" & text 2 thru 3 & "-" & text -4 thru -1
            end formatDate
    on addattachz(theFile)
        set theFile1 to theFile as string
        tell application "Keyboard Maestro Engine"
            setvariable "theFile" to theFile1
            delay 2
            do script "Run script to add attachments"
        end tell
    end addattachz
    on setthesubabo(subjectz, bodyz)
        tell application "Microsoft Outlook"
            try
            set theWin to window "<no subject> • jonathan@jerrydefalco.com"
            on error
            set theWin to front draft window
            end try
            set outgo to object of theWin
            set subject of outgo to subjectz
        end tell
        tell application "Keyboard Maestro Engine"
            setvariable "compvotext" to " " & bodyz as string
            do script "Add name to compvo email"
        end tell
    end setthesubabo
    on setthesub(subbyz)
        tell application "Microsoft Outlook"
            try
            set theWin to window "<no subject> • jonathan@jerrydefalco.com"
            on error
            set theWin to front draft window
            end try
            set outgo to object of theWin
            set subject of outgo to subbyz
        end tell
    end setthesub
    on setthestat(statusI, datalinee, typeC)
    #on setthestat(statusI, datalinee)
    log "setstatus func var statusI & typeC " & statusI & typeC
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
                    set rangz to "I" & datalinee
                    set value of cell rangz to statusI
                    set rangz to "D" & datalinee
                    set rangzx to "F" & datalinee
                    set tempcli to value of cell rangz
                    set temptit to value of cell rangzx
                    set findthesesst to temptit & tempcli
                    set findthesess to my theSplit(findthesesst, " ") as string
                end tell
            end tell
            #display dialog findthesess
            tell workbook "Audio Production Sheet.xlsx"
                tell sheet "Audio"
                    set newdataliz to {}
                    set searchydoi to 3
                    set datelz to 0
                    set typec1 to typeC
                    set failed2find to 0
                    set blankz to true
                    repeat while blankz is true
                        if typeC is "COMP VO"
                            if failed2find is 0 then
                                set typec1 to "60 radio"
                            else if failed2find is 1 then
                                set typec1 to "30 TV"
                            end if
                        else if typeC is "N/A"
                            if failed2find is 0 then
                                set typec1 to "60 radio"
                                else if failed2find is 1 then
                                set typec1 to "30 TV"
                            end if
                        end if
                    if failed2find is greater than 1 then
                    repeat 100 times
                        set rangt to "B" & searchydoi
                        set rangc to "A" & searchydoi
                        set ranzi to (value of cell rangt) & (value of cell rangc)
                        set strz to my theSplit(ranzi, " ") as string
                        try
                        set strz to my theSplit(strz, "(Jerry)") as string
                        end try
                        try
                        set strz to my theSplit(strz, "(Max)") as string
                        end try
                        #display dialog strz & " " & findthesess
                        if strz contains findthesess then
                            set datelz to searchydoi
                            exit repeat
                        end if
                    set searchydoi to searchydoi + 1
                    end repeat
                    else
                    repeat 100 times
                        set typrang to cell ("C" & searchydoi)
                        if (value of typrang) as string contains typec1 as string then
                        set rangt to "B" & searchydoi
                        set rangc to "A" & searchydoi
                        set ranzi to (value of cell rangt) & (value of cell rangc)
                        set strz to my theSplit(ranzi, " ") as string
                        try
                        set strz to my theSplit(strz, "(Jerry)") as string
                        end try
                        try
                        set strz to my theSplit(strz, "(Max)") as string
                        end try
                        #display dialog strz & " " & findthesess
                        if strz contains findthesess then
                            set datelz to searchydoi
                            set blankz to false
                            exit repeat
                        end if
                        end if
                    set searchydoi to searchydoi + 1
                    end repeat
                    set blankz to false
                    end if
                    if datelz is 0 then
                        failed2find = failed2find + 1
                    else
                    set rangz to "D" & datelz
                    set value of cell rangz to statusI
                    set rangzx to "A" & datelz & ":D" & datelz
                    if statusI contains "Parts in"
                        set bold of font object of range rangzx to true
                        set color of font object of range rangzx to {0, 0, 0}
                        set value of cell ("H" & datelz) to ""
                    else if statusI contains "sent"
                        set color of font object of range rangzx to {255, 0, 0}
                        set value of cell ("H" & datelz) to ""
                    else if statusI contains "sent"
                        set color of font object of range rangzx to {255, 0, 0}
                        set value of cell ("H" & datelz) to ""
                    else
                        set bold of font object of range rangzx to false
                        set color of font object of range rangzx to {0, 0, 0}
                        end if
                    end if
                    set searchydoi to datelz + 1
                    set end of newdataliz to datelz
                    end repeat
                end tell
            end tell
        end tell
        return(newdataliz)
    end setthestat
    on getsessiontermf(termyz)
        
                set subjem to termyz
        tell application "Microsoft Excel"
                tell workbook "Audio Production Sheet.xlsx"
                    tell sheet "Audio"
            set innt to 3
            set boolent to true
            repeat while boolent is true
                set myrgo to "A" & innt
                set myrgstr to value of range myrgo as string
                try
                    set tempyst to my theSplit (mygstr, "(Max)") as string
                end try
                try
                    set tempyst to my theSplit (mygstr, "(Jerry)") as string
                end try
                try
                set myrgstr to tempyst as string
                if subjem contains myrgstr then
                    set theinty to innt
                    set boolent to false
                    exit repeat
                else if innt is greater than 50 then
                    #display dialog "failed to find in production sheet"
                    exit repeat
                end if
                end try
                set innt to innt + 1
            end repeat
            if boolent is false then
            set myrangfo to "A" & theinty
            set myrangft to "B" & theinty
            set sessnlook to (value of range myrangft & " " & value of range myrangfo) as string
            try
            set sessnlook to my theSplit(sessnlook, "(Max)") as string
            set sessnlook to my theSplit(sessnlook, "(Jerry)") as string
            end try
            else
            set sessnlook to subjem
            end if
            try
            set sessnlook to my trimThis(sessnlook, true, "full")
            end try
            end tell
            end tell
                tell workbook "Database.xlsx"
                    tell sheet "Main Base"
            set innt to 2
            set boolent to true
            repeat while boolent is true
                set myrgo to "A" & innt
                if (value of range myrgo) as string contains (sessnlook as string) then
                    set dataline to innt
                    set boolent to false
                end if
                
                set innt to innt + 1
                
            end repeat
            set myrangf to "A" & dataline
            set sessnme to value of range myrangf as string
            set sessnme to my theSplit(sessnme, ".ptx") as string
            set outputz to {sessnme, dataline}
            end tell
            end tell
        end tell
        return (outputz)
    end getsessiontermf
on getsessiontermff(termyz)
        log "getsessiontermff"
        set perzalike to {}
        set intz12 to 1
        set neinty to 0
        set thisthing to true
        set theArray1 to {}
repeat while thisthing is true
    set getout1 to false
    set menuitemz to (my dynamicMenu)
    set listz to menuitemz's itemTitles() as list
    set intz1 to 1
    set countydo to count listz
        repeat with itemz1 in listz
        set end of perzalike to 0
        set newlistz1 to my theSplit(itemz1, " | ")
        if item 1 of newlistz1 contains "_C" then
        else
        set thestuffname to (item 1 of newlistz1 & " " & item 2 of newlistz1) as string
        #set newlistz2 to my theSplit(thestuffname, " ")
        set termylist to my theSplit(termyz, " ")
        set intz2 to 1
        log "getsessiontermff var thestuffname: " & thestuffname
        set theArstr to theArray1 as string
        repeat with the2item in termylist
            log "getsessiontermff var the2item & thestuffname: " & the2item & " " & thestuffname
            if the2item is "" then
            #else if the2item contains "mitsubishi"
            #else if the2item contains "kia"
            #else if the2item contains "Auto" then
            else
            if thestuffname contains the2item as string then
                set item intz12 of perzalike to (item intz12 of perzalike) + 1
                #set datalinee to my menutocellnum(itemz1)
                #if theArstr does not contain datalinee as string
                #set outputz1 to itemz1
                #set getout1 to true
                #end if
            end if
            end if
            set intz2 to intz2 + 1
            
        end repeat
        end if
        set intz12 to intz12 + 1
        #if getout1 is true
            #exit repeat
        #end if
        end repeat
        set intz1 to 1
        set greatest to 0
        repeat with stringvar1 in perzalike
            log "getsessiontermff var stringvar1: " & stringvar1
            if stringvar1 is greater than greatest then
                set outputz1 to item intz1 of listz as string
                set greatest to stringvar1
                log "getsessiontermff var outputz1: " & outputz1
            end if
        set intz1 to intz1 + 1
        end repeat
        try
            set datalinee to my menutocellnum(outputz1)
            set end of theArray1 to datalinee & " "
            tell application "Microsoft Excel"
                tell workbook "Database.xlsx"
                    tell sheet "Main Base"
                        set rangz to "F" & datalinee
                        set titn to value of cell rangz as string
                        set rangz to "D" & datalinee
                        set clin to value of cell rangz as string
                    end tell
                end tell
            end tell
        set outputz to {(titn & " " & clin), datalinee}
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
                    set theRanz to (cell ("G" & datalinee))
                    set theType to (value of theRanz) as string
                    #display dialog theType
                    if theType does not contain "30 TV"
                        if theType does not contain "60 RADIO"
                        set end of theArray1 to datalinee
                        else
                        set thisthing to false
                        exit repeat
                        end if
                    else
                    set thisthing to false
                    exit repeat
                    end if
                end tell
            end tell
        end tell
        on error
            set datalinee to "N/A"
            return {0,0}
        end try
        end repeat
        log "getsessiontermff var outputz1: " & outputz
        return (outputz)
    end getsessiontermff
    on theChekening(datalinee, dir1, sendername, namez, emailid, theMsg, sessnamez)
        set thisThingy to true
        repeat while thisThingy is true
        try
        log "theChekening started with var:" & datalinee & dir1 & sendername & namez
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
                    set rangz to "H" & datalinee
                    set docdir to value of cell rangz as string
                end tell
            end tell
        end tell
        try
            set docxHar to my theSplit(docdir, "*")
            repeat with doc1 in docxHar
                if doc1 contains "60" then
                    set docdir to doc1
                    exit repeat
                else if doc1 contains "TV"
                    if doc1 contains "30"
                        set docdir to doc1
                        exit repeat
                    end if
                end if
            end repeat
        end try
        log "theChekening var: " & datalinee as string
        log "theChekening var: " & docdir as string
        log "theChekening var: " & dir1 as string
        
        tell application "Finder"
        if docdir contains "/" then
            open docdir as posix file
        else
            open docdir
        end if
        if dir1 contains "/" then
            
            open dir1 as posix file
        else
            open (dir1 as string) as alias
        end if
        end tell
        display dialog "Got audio file: " & namez buttons {"It's good", "*Problem*", "Wrong session"}
        if the button returned of the result is "It's good" then
            set thisThingy to false
            set addpartinsh to true
            tell application "Microsoft Word" to quit
            tell application "QuickTime Player" to quit
            tell application "Microsoft Excel"
                tell workbook "Database.xlsx"
                    tell sheet "Main Base"
                        set ranzi to "K" & datalinee
                        set blez to value of cell ranzi as string
                        set blez to my trimThis(blez, true, "full")
                        set voica to my theSplit(blez, " ")
                        set voiceint to count voica
                        set int1 to 1
                        set rangz to "L" & datalinee
                        set value of cell rangz to (value of cell rangz) & (dir1 as string) & "*"
                        set allreieved to true
                        set alladir to value of cell rangz
                        repeat voiceint times
                            set voicat to item int1 of voica
                            if alladir contains voicat then
                            else
                            set allreieved to false
                            end if
                        set int1 to int1 + 1
                        end repeat
                        set ranzi to "p" & datalinee
                        set value of cell ranzi to value of cell ranzi & "*" & emailid
                        try
                        if allreieved is true then
                            set theranz to "C" & datalinee
                            set sessfoldc to value of cell theranz
                            set theranz to "L" & datalinee
                            set ranzi to "K" & datalinee
                            set value of cell theranz to (value of cell ranzi)
                            if sessfoldc is not ""
                                tell application "Finder"
                                set alladir to my trimThis(alladir, true, "full")
                                set allfilez to my theSplit(alladir, "*")
                                #display dialog allfilez as string
                                repeat with dir2 in allfilez
                                    set dir2str to dir2 as string
                                if dir2str contains "/" then
                                    try
                                    set commandydoo to do shell script "/usr/local/Cellar/tag/0.10_1/bin/tag -s red " & "'" & dir2 & "'"
                                    end try
                                    move (dir2 as posix file) to sessfoldc
                                    delete dir2 as posix file
                                else if dir2str contains ":" then
                                try
                                    set commandydoo to do shell script "/usr/local/Cellar/tag/0.10_1/bin/tag -s red " & "'" & posix path of (dir2) & "'"
                                end try
                                    move dir2 to sessfoldc
                                    delete dir2
                                else
                                end if
                                end repeat
                                end tell
                            else
                            display dialog "No session folder"
                            end if
                            set datelz to item 1 of my setthestat("Parts in house", datalinee, "N/A")
                            set addpartinsh to false
                        else
                        set addpartinsh to true
                        set datelz to item 1 of my setthestat("Waiting on parts", datalinee, "N/A")
                        end if
                        if allreieved is true then
                            set addpartinsh to false
                        end if
                        on error
                        end try
                    end tell
                end tell
                try
                set sendername to item 1 of sendername as string
                if addpartinsh is true then
                    tell workbook "Audio Production Sheet.xlsx"
                        tell sheet "Audio"
                               set rng1 to cell ("H" & datelz)
                               set value of rng1 to (value of rng1) as string & sendername & " is in, "
                        end tell
                    end tell
                else
                tell workbook "Audio Production Sheet.xlsx"
                    tell sheet "Audio"
                           set rng1 to cell ("H" & datelz)
                           set value of rng1 to ""
                    end tell
                end tell
                end if
                set faileded to "False"
                on error
                set faileded to "true"
                end try
            end tell
            
            log "theChekening var: " & dir2 as string
            log "thechecking var faileded: " & faileded

        else if the button returned of the result is "*Problem*" then
            tell application "QuickTime Player" to quit
            tell application "Microsoft Outlook" to set newMsg to reply to theMsg reply to all true
                tell application "Microsoft Outlook" to activate
                my addattachz(docdir)
        else if the button returned of the result is "Wrong session" then
        my ctwin1's makeKeyAndOrderFront_(1)
        end if
    on error
        log "Error"
    end try
    end repeat
    end theChekening
    on filzexists(theFolder, theFile)
    set rpdo to true
    set int1 to 1
    set theTfile to theFile
    repeat while rpdo is true
        set checky to false
        tell application "Finder"
            set thefiz to " "
            set dir3 to (theFolder & theTFile) as string
            try
            set thefiz to name of file dir3 as string
            end try
            if thefiz is " " then
                set theFile to theTFile
                set rpdo to false
                exit repeat
            else
            set theTFile to int1 & " " & theFile
            end if
        end tell
        set int1 to int1 + 1
        end repeat
    return(theFile)
    end filzexists
    on aprjeT_(sender)
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Approved Radio Jerry.emltpl" as Posix file)
    end aprjeT_
    on aprmaT_(sender)
    tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Approved Radio Max.emltpl" as Posix file)
    end aprmaT_
    on mixjeT_(sender)
    tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Jerry mix.emltpl" as Posix file)
    end mixjeT_
    on mixmaT_(sender)
    tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Max mix.emltpl" as Posix file)
    end mixmaT_
    on compjeT_(sender)
    tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/JERRY COMPVO.emltpl" as Posix file)
    end compjeT_
    on compmaT_(sender)
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Max COMPVO.emltpl" as Posix file)
    end compmaT_
    on aptvaT_(sender)
    tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Approved TV audio.emltpl" as Posix file)
    end aptvaT_
    on comptoT_(sender)
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/TOM COMPVO.emltpl" as Posix file)
    end comptoT_
on Newscriptread(listchecked)
log "Newscriptread"
    tell application "Microsoft Outlook"
        if count of listchecked is greater than 20
            my scriptpt2(listchecked)
            tell application "Microsoft Outlook" to set is read of listchecked to true
        else if count of listchecked is 1
        my scriptpt2(listchecked)
        tell application "Microsoft Outlook" to set is read of listchecked to true
        else
        set thecounty to count listchecked
        set some1 to thecounty
        repeat thecounty times
            set theMsg to item some1 of listchecked
            my scriptpt2(theMsg)
            set some1 to some1 - 1
            tell application "Microsoft Outlook" to set is read of theMsg to true
        end repeat
        end if
    end tell
end Newscriptread
on scriptpt2(theMsg)
tell application "Microsoft Outlook"
    set docxH to {}
    set docxHH to {}
    set trything to true
    set numofdocx to 0
    repeat while trything is true
    try
    set theSubject to subject of theMsg as string
    set thecontent to plain text content of theMsg as string
    exit repeat
    end try
    end repeat
    if thecontent contains "cut" then
        set cutdown to true
    else
        set cutdown to false
    end if
    if thecontent contains "on screen" then
        if thecontent contains "only" then
            set onscreenchangeonly to true
        else
            set onscreenchangeonly to false
        end if
        else
        set onscreenchangeonly to false
    end if
    delay 2
        set NstoreFolder to "LaCie:current scripts :completed scripts" as alias
        #tell application "Finder"
        set selection to mail folder "JP Scripts" of exchange account "jonathan@jerrydefalco.com"
        delay 3
        set selection to theMsg
        delay 3
        set msg to first item of (get current messages)
            #set sandboxDocumentFolder to (path to documents folder as string)
            #set sandboxAttachmentsFolder to sandboxDocumentFolder & "Attachments"
            set storeFolder to "Macintosh HD:Users:jonathanstoff:Downloads" as alias
            set allAttachments to attachments of msg
            repeat with thisAttachment in allAttachments
                set saveName to name of thisAttachment
                #display dialog saveName as string
                #set destName to sandboxAttachmentsFolder & ":" & saveName
                save thisAttachment in storeFolder
            end repeat
        #end tell
end tell
tell application "Finder"
    set docxHH to get every item of storeFolder
    
    set NstoreFolder to "LaCie:current scripts :completed scripts" as alias
repeat with docxdir in docxHH
    set tsn to name of docxdir as string
    if tsn contains "graphic" then
        delete  docxdir
    else if tsn contains ".doc" then
        set end of docxH to ((NstoreFolder as string) & tsn)
        move docxdir to NstoreFolder with replacing
        delete docxdir
    else
        delete docxdir
    end if
end repeat
end tell
    repeat with theDocxf in docxH
        set sessnameA to ""
        set clientnD to "" #yes
        set clientcE to "" #yes
        set titlenF to "" #yes
        set typesG to {} #yes
        set docxpathH to "" #yes
        set vaK to {} #yes
        set clientname to ""
        set titlename to ""
        set newcellnum to 0
        
        tell application "Finder" to open theDocxf
            delay 3
        tell application "Microsoft Word"
            set theDoc to name of active document as string
            if theDoc contains "Spanish" then
                set Spanish1 to true
            else
                set Spanish1 to false
            end if
                set theWin to name of front window
                set docxpathH to theDocxf
                set theSelectionn to selection
                tell selection
                    delay 1
                    set selection start to 2
                    set selection end to 2000
                    delay 2
                    set document1 to content as string
                        if content does not contain "TITLE:" then
                            set istv to true
                            set lines1 to my theSplit(document1, ":")
                            set line1 to item 1 of lines1
                            set line2 to item 2 of lines1
                            set lines2 to my theSplit(line2, "“")
                            try
                                set line3 to item 3 of lines2
                            on error
                                set line3 to item 2 of my theSplit(document1, "ANN")
                                set line3 to item 1 of my theSplit(line3, return)
                            end try
                            set clientname to line1
                            #set ttitname to my theSplit(item 2 of lines2, "”")
                            set ttitname to item 2 of lines2 as string
                            set typesG to item 1 of lines2
                            set titlename to my trimThis(ttitname, true, "full")
                            if titlename contains "REV" then
                                set fulltit to titlename
                                set titlenamedd to my theSplit(titlename, "REV")
                                set titlenamezzx to item 1 of titlenamedd
                                set titlenamezzz to my theSplit(titlenamezzx, "CON")
                                set titlename to item 1 of titlenamezzz
                                set titlename to my trimThis(titlename, true, "full")
                                set revision1 to true
                            else
                                set fulltit to titlename
                                set revision1 to false
                            end if
                            try
                                if titlename contains "’" then
                                    set titlename to my theSplit(titlename, "’") as string
                                end if
                            end try
                            set titlename to my trimThis(titlename, true, "full")
                            set clientname to my trimThis(clientname, true, "full")
                            set typesG to my trimThis(typesG, true, "full")
                            set titlenF to titlename
                        else
                        #not tv
                        set lines1 to my theSplit(document1, "TITLE:")
                        set line1 to item 1 of lines1
                        set line2 to item 2 of lines1
                        if line2 contains "NEED" then
                            set lines2 to my theSplit(line2, "NEED")
                        else if line2 contains "AIR" then
                            set lines2 to my theSplit(line2, "AIR")
                        end if
                        set line2 to item 1 of lines2
                        set line3 to item 2 of lines2
                        set lines3 to my theSplit(line3, "MUSIC:")
                        set line3 to item 1 of lines3
                        set line1array to my theSplit(line1, "spot:")
                        set clientname2 to item 1 of line1array as text
                        set spottype to item 2 of line1array as text
                        set clientname to item 1 of line1array as text
                        set line2array to line2 as string
                        set titlename to line2array
                        if titlename contains "REV" then
                            set fulltit to titlename
                            set titlenamedd to my theSplit(titlename, "REV")
                            set titlename to item 1 of titlenamedd
                            set titlename to my trimThis(titlename, true, "full")
                            set revision1 to true
                        else
                            set fulltit to titlename
                            set revision1 to false
                        end if
                        try
                            if titlename contains "’" then
                                set titlename to my theSplit(titlename, "’") as string
                                set fulltit to my theSplit(fulltit, "’") as string
                            end if
                        end try
                        set titlename to my trimThis(titlename, true, "full")
                        set fulltit to my trimThis(fulltit, true, "full")
                        set clientname to my trimThis(clientname, true, "full")
                        set spottype to my trimThis(spottype, true, "full")
                        if clientname contains ":" then
                            set clientname1 to my theSplit(clientname, ":")
                            set inttyc to (count of clientname1) - 1
                            set ehh to 2
                            set clientnamezx to {}
                            repeat inttyc times
                                if item ehh of clientname1 is not "" then
                                    set end of clientnamezx to item ehh of clientname1
                                end if
                                
                            end repeat
                            set clientname to clientnamezx as string
                            set clientname to my trimThis(clientname, true, "full")
                        end if
                        
                        set titlenF to fulltit
                        set typesG to spottype as string
                    end if
                end tell
                
                    if line3 contains "mark" then
                        if line3 contains "mark b"
                            set end of vaK to "Mark B "
                        else
                            set end of vaK to "Mark "
                        end if
                    end if
                    if line3 contains "jim" then
                        set end of vaK to "Jim "
                    end if
                    if line3 contains "don" then
                        set end of vaK to "Donovan "
                    end if
                    if line3 contains "ben" then
                        set end of vaK to "Ben "
                    end if
                    if line3 contains "brent" then
                        set end of vaK to "Brent "
                    end if
                    if line3 contains "david" then
                        set end of vaK to "David "
                    end if
                    if line3 contains "doak" then
                        set end of vaK to "Doak "
                    end if
                    if line3 contains "melissa" then
                        set end of vaK to "Melissa "
                    end if
                    if line3 contains "mike" then
                        set end of vaK to "Mike o "
                    end if
                    if line3 contains "rachel" then
                        set end of vaK to "Rachel "
                    end if
                    if line3 contains "sandra" then
                        set end of vaK to "Sandra "
                    end if
                    if line3 contains "WILNETTE" then
                        set end of vaK to "Wilnette"
                    else
                        if line3 contains "spanish" then
                            set vaK to "Spanish"
                        end if
                        if line2 contains "spn" then
                            set vaK to "Spanish"
                        end if
                        if vaK contains "Spanish" then
                            else
                            if line2 contains "spanish" then
                                set vaK to "Spanish"
                            end if
                        end if
                    end if
                    try
                    if titlenF contains "ann" then
                        set titlenF to item 1 of my theSplit(titlenF, "ANNOU")
                        set titlenF to trimthis(titlenF, true, "full")
                    end if
                    end try
                    if clientname contains "BENNET" then
                        set clientname to "BENNETTSVILLE HONDA"
                    end if
                    set clientnD to clientname
                    
                    set jerrysclientsl to my readFile("/Volumes/LaCie/Work/jclients.txt") as string
                    set tclientnD to clientnD
                    if clientnD contains "MATHEW" then
                        set tclientnD to "MATHEWS AUTO GROUP"
                    else if clientnD contains "BENZ" then
                        set tclientnD to "MERCEDES BENZ OF FT WAYNE"
                    else if clientnD contains "OPELIKA" then
                        set tclientnD to "OPELIKA FORD CDJR"
                    else if clientnD contains "STRONG" then
                        set tclientnD to "STRONG VW"
                    else if clientnD contains "PARK" then
                        set tclientnD to "AUTO WORLD MITSUBISHI"
                    else if clientnD contains "CAPE" then
                        set tclientnD to "CAPE CORAL KIA"
                    else if clientnD contains "BENNET" then
                        set tclientnD to "BENNETTSVILLE FORD/HONDA"
                    end if
                    if jerrysclientsl contains tclientnD then
                        set clientcE to "Jerry's"
                    else
                        set clientcE to "Jerry's"
                    end if
                    if theSubject contains "Jerry" then
                        set clientcE to "Jerry's"
                    else if theSubject contains "Max" then
                        set clientcE to "Jerry's"
                    end if
            if vaK contains "Wilnette" then
                else
                if vaK contains "Spanish" then
                else
                if onscreenchangeonly is false
                if (my printer) is true then
                   print active document
                end if
                end if
                end if
            end if
            
            quit
        end tell
            try
                set clientnD to my change_cased(clientnD)
                set temptitlenff to my theSplit(titlename, "“")
                set titlenF to temptitlenff as string
                set temptitlenff to my theSplit(titlename, "”")
                set titlenF to temptitlenff as string
                set temptitlenff to my theSplit(fulltit, "”")
                set fulltit to temptitlenff as string
                set temptitlenff to my theSplit(fulltit, "“")
                set fulltit to temptitlenff as string
            end try
            set termy to titlenF & " " & clientnD
            #set termy to my theSplit(termy, " ") as string
            set cellnumber to my scriptsearch1(termy, clientnD, titlenF)
            delay 1
            set datedue to my createdate(theSubject)
            set sec30bool to false
            if typesG contains "30" then
                if typesG contains "TV" then
                else
                set sec30bool to true
                end if
            end if
            tell application "Microsoft Excel"
                tell workbook "Database.xlsx"
                        tell sheet "Main Base"
                    set thedaytoday1ar to current date
                    set thedaytoday2 to (month of thedaytoday1ar & "/" & day of thedaytoday1ar) as string
                    set mygang to cell ("O" & cellnumber)
                    set value of mygang to thedaytoday2
                    set typesG to my change_cased(typesG)
                    set myrange to "D" & cellnumber
                    set value of cell myrange to clientnD as string
                    set myrange to "J" & cellnumber
                    set value of cell myrange to datedue as string
                    set myrange to "E" & cellnumber
                    set value of cell myrange to clientcE as string
                    set myrange to "F" & cellnumber
                    set value of cell myrange to fulltit as string
                    set myrange to "G" & cellnumber
                    if (value of cell myrange as string) contains (typesG as string) then
                    else
                        if value of cell myrange as string is "" then
                            set value of cell myrange to typesG as string
                        else
                            set value of cell myrange to value of cell myrange & "*" & typesG as string
                        end if
                    end if
                    set myrange to "H" & cellnumber
                    if value of cell myrange is "" then
                        set value of cell myrange to docxpathH as string
                    else
                        set value of cell myrange to value of cell myrange & "*" & docxpathH as string
                    end if
                    set myrange to "K" & cellnumber
                    set value of cell myrange to vaK as string
                    set myrange to "I" & cellnumber
                    if cutdown is false then
                    set value of cell myrange to "Waiting on Parts" as string
                    else
                    set value of cell myrange to "Parts in house" as string
                    end if
                        end tell
                end tell
                if vaK contains "WILNETTE" then
                    if onscreenchangeonly is false
                        my addselectoashz(cellnumber, typesG)
                    end if
                else
                if onscreenchangeonly is false
                    if cutdown is false then
                        if sec30bool is false then
                        if revision1 is true then
                            my sendvoREVemail(vaK, clientnD, Spanish1, datedue, docxpathH)
                        else
                            my sendvoemail(vaK, clientnD, Spanish1, datedue, docxpathH)
                        end if
                        end if
                    end if
                    my addselectoashz(cellnumber, typesG)
                end if
                end if
                my adddatalinetocurrentz(cellnumber)
            end tell
    end repeat
end scriptpt2
on scriptsearch1(searchterm, clin, titn)
set searchterm to my theSplit(searchterm, "“") as string
set clin to my theSplit(clin, "“") as string
set titn to my theSplit(titn, "“") as string
log "scriptsearch1 all var: " & searchterm & clin & titn
    tell application "Microsoft Excel"
               tell workbook "Database.xlsx"
                   tell sheet "Main Base"
                       set theadderar to {}
                       set bzz to true
                       set rangz to ""
                       set tcnumin to 2
                       set myRangez to range "A:O"
                       set cellznum to find myRangez what "*" after (cell 1 of myRangez) look in values search direction search previous
                       set rangz1 to ((first row index of cellznum) + 1)
                    try
                       set rangez to find range ("A" & tcnumin & ":H" & 2000) what searchterm
                       set tcnum to (first row index of rangez)
                       set rangz1 to tcnum
                       on error
                        
                   end try
                        return(rangz1)
                   end tell
               end tell
           end tell
    log "scriptsearch1 ended"
end scriptsearch1
    on adddatalinetocurrentz(cellnumber)
        set tempar1 to {}
        tell application "Microsoft Excel"
            tell workbook "Database.xlsx"
                tell sheet "Main Base"
                    set tempran1 to "F" & cellnumber
                    set tempran2 to "D" & cellnumber
                    set tempran3 to "E" & cellnumber
                    set tempstr1 to value of cell tempran1 & " | " & value of cell tempran2 & " | " & value of cell tempran3 & " | " & "<" & cellnumber & ">"
                end tell
            end tell
        end tell
        my writetopsess(tempstr1)
    end adddatalinetocurrentz
on addselectoashz(cellnumber, typeC)
set dataline to cellnumber
            tell application "Microsoft Excel"
                tell workbook "Database.xlsx"
                    tell sheet "Main Base"
                        set tempran1 to "F" & cellnumber
                        set titB to value of cell tempran1
                        set tempran2 to "D" & cellnumber
                        set cliA to value of cell tempran2
                        set tempran3 to "E" & cellnumber
                        set clicon to value of cell tempran3
                        set rangx to "I" & dataline
                        set statusD to value of cell rangx
                        
                        set rangx to "O" & dataline
                        set SubdE to value of cell rangx
                        set rangx to "J" & dataline
                        set NeedbF to value of cell rangx
                        set rangx to "K" & dataline
                        set Anncr to value of cell rangx
                    end tell
                end tell
                if clicon contains "Jerry" then
                    set cliA to cliA & "(Jerry)"
                else if clicon contains "Max" then
                    set cliA to cliA & "(Max)"
                end if
                set Anncrsh to {}
                if Anncr contains "Mark" then
                    if Anncr contains "Mark B" then
                        set end of Anncrsh to "MB "
                    else
                        set end of Anncrsh to "MM "
                    end if
                end if
                    if Anncr contains "Rachel" then
                        set end of Anncrsh to "RB "
                    end if
                    if Anncr contains "Jim" then
                        set end of Anncrsh to "JM "
                    end if
                    if Anncr contains "Chris" then
                        set end of Anncrsh to "CC "
                    end if
                    if Anncr contains "Melissa" then
                        set end of Anncrsh to "MEL "
                    end if
                    if Anncr contains "Mike O" then
                        set end of Anncrsh to "MO "
                    end if
                    if Anncr contains "Donovan" then
                        set end of Anncrsh to "DV "
                    end if
                    if Anncr contains "Brent" then
                        set end of Anncrsh to "BM "
                    end if
                    if Anncr contains "Andrea" then
                        set end of Anncrsh to "AB "
                    end if
                    if Anncr contains "Ben" then
                        set end of Anncrsh to "BB "
                    end if
                    if Anncr contains "David" then
                        set end of Anncrsh to "DT "
                    end if
                    if Anncr contains "Mitch" then
                        set end of Anncrsh to "MP "
                    end if
                    if Anncr contains "Paco" then
                        set end of Anncrsh to "PL "
                    end if
                    if Anncr contains "Sandra" then
                        set end of Anncrsh to "SS "
                    end if
                    if Anncr contains "Doak" then
                        set end of Anncrsh to "DB "
                    end if
                    if Anncr contains "Rob" then
                        set end of Anncrsh to "RM "
                    end if
                tell workbook "Audio Production Sheet.xlsx"
                    tell sheet "Audio"
                        set int2 to 3
                        repeat 50 times
                            set ran3 to "B" & int2
                            set tsessname to value of cell ran3
                            if tsessname is "" then
                                set ndata to int2
                                exit repeat
                            end if
                            set int2 to int2 + 1
                        end repeat
                        set rangzx to "A" & ndata & ":F" & ndata
                        if statusD contains "Parts in"
                            set bold of font object of range rangzx to true
                            set color of font object of range rangzx to {0, 0, 0}
                        else
                            set bold of font object of range rangzx to false
                            set color of font object of range rangzx to {0, 0, 0}
                        end if
                        set ran3 to "A" & ndata
                        set value of cell ran3 to cliA
                        set ran3 to "B" & ndata
                        set value of cell ran3 to titB
                        set ran3 to "C" & ndata
                        set value of cell ran3 to typeC
                        set ran3 to "D" & ndata
                        set value of cell ran3 to statusD
                        set ran3 to "E" & ndata
                        set value of cell ran3 to SubdE
                        set ran3 to "F" & ndata
                        set value of cell ran3 to NeedbF
                        set ran3 to "G" & ndata
                        set value of cell ran3 to Anncrsh as string
                    end tell
                end tell
            end tell
        end addselectoashz
        on list2string(theList, theDelimiter)
            set theBackup to AppleScript's text item delimiters
            set AppleScript's text item delimiters to theDelimiter
            set theString to theList as string
            set AppleScript's text item delimiters to theBackup
            return theString
        end list2string
        on remove:remove_string fromString:source_string
            set s_String to NSString's stringWithString:source_string
            set r_String to NSString's stringWithString:remove_string
            return s_String's stringByReplacingOccurrencesOfString:r_String withString:""
        end remove:fromString:
        property NSString : a reference to current application's NSString
        set saveTID to text item delimiters
    on createdate(subjectline)
        
        set duedate to ""
        if subjectline contains "asap" then
            set duedate to "ASAP"
        end if
        if subjectline contains "ASAP" then
        else
            set subject1array to my theSplit(subjectline, " ")
            set arraycount to count items in subject1array
            #display dialog arraycount as string
            set repeatcount to arraycount
            repeat arraycount times
                if item repeatcount of subject1array contains "/" then
                    set date1 to item repeatcount of subject1array
                    set lengthofdate1 to length of date1
                    if lengthofdate1 is greater than 5 then
                        set date1array to my theSplit(date1, "/")
                        set duedate to item 1 of date1array & "/" & item 2 of date1array
                    else
                        set duedate to date1
                    end if
                end if
                if item repeatcount of subject1array contains (92) then
                    set lengthofdate1 to length of date1
                    set date1 to item repeatcount of subject1array
                    if lengthofdate1 is greater than 5 then
                        set date1array to my theSplit(date1, (92))
                        set duedate to item 1 of date1array & "/" & item 2 of date1array
                    else
                        set duedate to date1
                    end if
                end if
                set repeatcount to repeatcount - 1
            end repeat
        end if
        if duedate is "" then
            set duedate to "ASAP"
        end if
        return(duedate)
    end createdate
on sendvoemail(vaK, clientnD, Spanish1, datedue, docxpathH)
    #display dialog vaK
    log "started sendvoemail"
    log "sendvoemail var " & vaK
    set vaK to vaK as string
    if vaK contains "Mark" as string then
        if vaK contains "Mark B" as string then
            tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/MarkB SCRIPT.emltpl" as posix file)
                delay 0.7
                my createthemail(clientnD, datedue, docxpathH)
        else
        if datedue does not contain "asa"
        set datedue to my trimThis(datedue, true, "full")
                set dueardate to my theSplit(datedue, "/")
                set theDayt to (item 2 of dueardate) as number
                set theMontht to (item 1 of dueardate) as number
                if theDayt is 1 then
                    if theMontht is 1 then
                        set ntheMontht to 12
                    else
                        set ntheMontht to theMontht - 1
                    end if
                    if ntheMontht is 2 then
                        set theDay to 28
                    else if ntheMontht is 4
                        set theDay to 30
                    else if ntheMontht is 6
                        set theDay to 30
                    else if ntheMontht is 9
                        set theDay to 30
                    else if ntheMontht is 11
                        set theDay to 30
                    else
                        set theDay to 31
                    end if
                else
                    set ntheMontht to theMontht
                set theDay to theDayt - 1
                
            end if
            set thedatedue to (ntheMontht & "/" & theDay) as string
            else
            set thedatedue to datedue
            end if
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Mark SCRIPT.emltpl" as posix file)
            delay 0.7
            my createthemail(clientnD, thedatedue, docxpathH)
        end if
    end if
    if vaK contains "Paco" as string then
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/paco with disclaimer.emltpl" as posix file)
            delay 0.7
            my createthemail(clientnD, datedue, docxpathH)
    end if
    if vaK contains "Jim" as string then
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Jimm SCRIPT.emltpl" as posix file)
            delay 0.7
            my createthemail(clientnD, datedue, docxpathH)
    end if
    if vaK contains "Donovan" as string then
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Don script.emltpl" as posix file)
            delay 0.7
            my createthemail(clientnD, datedue, docxpathH)
    end if
    if vaK contains "Brent" as string then
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/brent script.emltpl" as posix file)
            delay 0.7
            my createthemail(clientnD, datedue, docxpathH)
    end if
    if vaK contains "David" as string then
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/david t script.emltpl" as posix file)
            delay 0.7
            my createthemail(clientnD, datedue, docxpathH)
    end if
    if vaK contains "Doak" as string then
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Doak Script.emltpl" as posix file)
            delay 0.7
            my createthemail(clientnD, datedue, docxpathH)
    end if
    if vaK contains "Melissa" as string then
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/mel script.emltpl" as posix file)
            delay 0.7
            my createthemail(clientnD, datedue, docxpathH)
    end if
    if vaK contains "Mike" as string then
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Mike o.emltpl" as posix file)
            delay 0.7
            my createthemail(clientnD, datedue, docxpathH)
    end if
    if vaK contains "Rachel" as string then
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Rachel SCRIPT no dis.emltpl" as posix file)
            delay 0.7
            my createthemail(clientnD, datedue, docxpathH)
    end if
    if vaK contains "Sandra" as string then
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/sandra script.emltpl" as posix file)
            delay 0.7
            my createthemail(clientnD, datedue, docxpathH)
    end if
    if vaK contains "Andrea" as string then
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/Andrea withdis.emltpl" as posix file)
            delay 0.7
            my createthemail(clientnD, datedue, docxpathH)
    end if
    if Spanish1 is true then
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/translate andrea.emltpl" as posix file)
            delay 0.7
            my createthemailt(clientnD, datedue, docxpathH)
    end if
end sendvoemail
on sendvoREVemail(vaK, clientnD, Spanish1, datedue, docxpathH)
set vaK to vaK as string
    if vaK contains "Mark" as string then
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/revisions/Mark SCRIPT.emltpl" as posix file)
            delay 0.7
            my createthemail(clientnD, datedue, docxpathH)
    end if
    if vaK contains "Paco" as string then
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/revisions/paco with disclaimer.emltpl" as posix file)
            delay 0.7
            my createthemail(clientnD, datedue, docxpathH)
    end if
    if vaK contains "brent" as string then
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/revisions/brent script.emltpl" as posix file)
            delay 0.7
            my createthemail(clientnD, datedue, docxpathH)
    end if
    if vaK contains "Jim" as string then
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/revisions/jim script rev.emltpl" as posix file)
            delay 0.7
            my createthemail(clientnD, datedue, docxpathH)
    end if
    if vaK contains "david" as string then
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/revisions/david t script.emltpl" as posix file)
            delay 0.7
            my createthemail(clientnD, datedue, docxpathH)
    end if
    if vaK contains "doak" as string then
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/revisions/Doak Script.emltpl" as posix file)
            delay 0.7
            my createthemail(clientnD, datedue, docxpathH)
    end if
    if vaK contains "melissa" as string then
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/revisions/mel script.emltpl" as posix file)
            delay 0.7
            my createthemail(clientnD, datedue, docxpathH)
    end if
    if vaK contains "mike" as string then
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/revisions/Mike o.emltpl" as posix file)
            delay 0.7
            my createthemail(clientnD, datedue, docxpathH)
    end if
    if vaK contains "rachel" as string then
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/revisions/Rachel SCRIPT.emltpl" as posix file)
            delay 0.7
            my createthemail(clientnD, datedue, docxpathH)
    end if
    if vaK contains "sandra" as string then
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/revisions/sandra script.emltpl" as posix file)
            delay 0.7
            my createthemail(clientnD, datedue, docxpathH)
    end if
    if Spanish1 is true then
        tell application "Finder" to open ("/Users/jonathanstoff/Documents/email templates/revisions/translate andrea.emltpl" as posix file)
            delay 0.7
            my createthemailt(clientnD, datedue, docxpathH)
    end if
end sendvoemail
on createthemail(clientnD, datedue, docxpathH)
log "createthemail"
set thenewsubby to clientnD & " Need " & datedue as string
    tell application "Keyboard Maestro Engine"
        do script "63DDBEE7-06DA-4C17-BCFF-4F7D0114E744"
    end tell

    my setthesubabo(thenewsubby, datedue)
    my addattachz(docxpathH)
    delay 2
    tell application "Keyboard Maestro Engine"
        do script "914B9E1F-4301-461C-850D-EB805DCEC47C"
    end tell
end createthemail
on createthemailt(clientnD, datedue, docxpathH)
log "createthemailt"
    set thenewsubby to clientnD & " Translation" as string
    tell application "Keyboard Maestro Engine"
        do script "63DDBEE7-06DA-4C17-BCFF-4F7D0114E744"
    end tell
        my setthesub(thenewsubby)
        my addattachz(docxpathH)
        delay 2
        tell application "Keyboard Maestro Engine"
            do script "914B9E1F-4301-461C-850D-EB805DCEC47C"
        end tell
    end createthemailt
end script

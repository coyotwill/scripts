tell application "Microsoft Outlook"
	-- Bring app to FG if it was already open.
	activate
	if exists window "Calendar" then
		-- Unminimize, open default window
		reopen
		-- Set Calendar to foremost
		set index of window "Calendar" to 1
	else
		tell application "System Events"
			-- Open new window, then navigate to calendar view
			keystroke "n" using {command down, option down}
			keystroke 2 using {command down}
		end tell
		with timeout of 2 seconds
			repeat until exists (window "Calendar")
				delay 0.1
			end repeat
			set bounds of window "Calendar" to {50, 80, 1500, 1000}
		end timeout
	end if
end tell

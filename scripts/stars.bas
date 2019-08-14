sub stars(n)
	if n > 0 then 
		cs.system.console.write "*" 
		stars n-1
	end if
end sub
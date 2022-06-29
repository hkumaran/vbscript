Dim a
a="I am sad 2 3"
b=split(a," ")
for i=lbound(b) to ubound(b)
	msgbox b(i)
	if instr(1,b(i),"sa") then
		exit for
	end if
Next
		
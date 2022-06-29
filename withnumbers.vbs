Dim a
sum=0
a="I am happy 2 3"
b=split(a," ")
for i=lbound(b) to ubound(b)
	
	if isnumeric(b(i))=true then
		sum=sum+b(i)
		
	end if
msgbox sum
Next
		
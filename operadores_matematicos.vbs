dim n1, n2, soma, op, resp, qtde, sair
call operacao
sub operacao()
randomize(second(time))
n1=cint(int(rnd * 10) + 1)
n2=cint(int(rnd * 10) + 1)
op=int(rnd * 3) + 1
Select case op
	case 1
		soma=cint(n1+n2)
		op = "+"
	case 2
		soma=cint(n1-n2)
		op = "-"
	case 3
		soma=cint(n1*n2)
		op = "x"
	case else
		call operacao
end Select
resp=cint(inputbox("Resolva a operaçao matematica" + vbnewline &_
			       "Resolva : "& n1 &" "& op &" "& n2 &"","Operadores"))
if resp=soma then
	qtde = qtde + 1
	msgbox("Voce acertou!!!" + vbnewline &_
	       "Qtde de acertos "& qtde &""),vbinformation+vbokonly,"Operadores Matematicos"
	sair=msgbox("Deseja sair???",vbquestion+vbyesno,"AVISO")
		if sair=vbyes then
			wscript.quit
		else
			call operacao
		end if
else 
	msgbox("Voce Errou!!!" + vbnewline &_
	       "Qtde de acertos "& qtde &""),vbinformation+vbokonly,"Operadores Matematicos"
	sair=msgbox("Deseja sair???",vbquestion+vbyesno,"AVISO")
		if sair=vbyes then
			wscript.quit
		else
			call operacao			
		end if   
end if
end sub	
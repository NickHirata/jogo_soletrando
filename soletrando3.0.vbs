dim nivel(40),qtde, n, i, nome
dim audio,pontos,resp,valor,situacao,qtde1

call jogador

sub jogador()
nome=inputbox("Digite seu nome: ")
pontos=0
qtde=0
qtde1=0
call aviso
end sub

sub aviso()
msgbox("===================================" + vbNewLine &_
	   "Use somente letras minusculas " + vbNewLine &_
	   "e nao coloque acento, " + vbNewLine &_
	   "pois conta como erro." + vbNewLine &_
	   "===================================" + vbNewLine &_
	   "REGRAS:" + vbNewLine &_
	   "A cada 5 acertos consecutivos muda o nivel;" + vbNewLine &_
	   "Caso erre, os pontos serao zerados." + vbNewLine &_
	   "===================================" + vbNewLine &_
	   "Bom jogo")
call palavras
end sub

sub carregar_voz()

set audio= CreateObject("SAPI.SPVOICE")
audio.volume=100
audio.rate=2 
		audio.speak(nivel(n))
call resposta
end sub

sub resposta() 


resp=inputbox("Digite a palavra ouvida" + vbNewLine &_
			  "===================================" + vbNewLine &_
			  "[O]uvir a palavra novamente" + vbNewLine &_
			  "[P]ular a palavra" + vbNewLine &_
			  "[S]air do jogo" + vbNewLine &_
			  "===================================")
if lcase(resp) = nivel(n) Then
	qtde=qtde+1
	pontos=valor+pontos
	call acerto
elseif lcase(resp) = "o" then
	call carregar_voz
elseif lcase(resp)="s" then 
	wscript.quit
elseif lcase(resp) = "p" Then
	call palavras
else
	qtde1=qtde1+1
	if situacao = "muito facil" then
		pontos = 0
	elseif situacao = "facil" then
		pontos = 50
	elseif situacao = "intermediario" then
		pontos = 550
	else 
		pontos = 5550
	end if
	call erro
end if
end sub

sub palavras()
if pontos >=0 and pontos <50 then
	situacao="muito facil"
	valor = 10
    nivel(0)="amor"
	nivel(1)="paz"
	nivel(2)="ano"
	nivel(3)="mundo"
	nivel(4)="asa"
	nivel(5)="rua"
	nivel(6)="carro"
	nivel(7)="vida"
	nivel(8)="sutil"
	nivel(9)="luz"
	Randomize(second(time))
n=int(rnd * 9)
call carregar_voz
elseif pontos >=50 and pontos <550 then
	situacao="facil"
	valor = 100
	nivel(0)="programas"
	nivel(1)="computador"
	nivel(2)="cluster"
	nivel(3)="arquitetura"
	nivel(4)="televisao"
	nivel(5)="monitor"
	nivel(6)="faculdade"
	nivel(7)="tecnologia"
	nivel(8)="engenharia"
	nivel(9)="software"
	Randomize(second(time))
n=int(rnd * 9)
call carregar_voz
elseif pontos >=550 and pontos <5550 then
	situacao="intermediario"
	valor = 1000
	nivel(0)="propolis"
	nivel(1)="fuzarca"
	nivel(2)="chapoletada"
	nivel(3)="yorkshire"
	nivel(4)="palimpsesto"
	nivel(5)="imprescindivel"
	nivel(6)="inexoravel"
	nivel(7)="tertulia"
	nivel(8)="inconstitucionalissimamente"
	nivel(9)="tergiversar"
	Randomize(second(time))
n=int(rnd * 9)
call carregar_voz
elseif pontos >=5550 then
	situacao="dificil"
	valor = 1000000
	nivel(0)="piperidinoetoxicarbometoxibenzofenona"
	nivel(1)="paraclorobenzilpirrolidinonetilbenzimidazol"
	nivel(2)="pneumartrorradiografia"
	nivel(3)="anticonstitucionalissimamente"
	nivel(4)="dacriocistossiringotomia"
	nivel(5)="pneumoultramicroscopicossilicovulcanoconiotico"
	nivel(6)="aeropiesotermoterapico"
	nivel(7)="fotocromometalografico"
	nivel(8)="histerossalpingectomia"
	nivel(9)="ooforossalpingectomia"
	Randomize(second(time))
n=int(rnd * 9)
call carregar_voz
end if

	
end sub




sub acerto()
msgbox("-------------------------" + vbNewLine &_
	   "Nivel: "& situacao &"" + vbNewLine &_
	   "-------------------------" + vbNewLine &_
	   "Voce acertou "& nome &"" + vbNewLine &_
	   "-------------------------" + vbNewLine &_
	   "Erros: "& qtde1 &"" + vbNewLine &_
	   "Acertos: "& qtde &"" + vbNewLine &_
	   "Pontos: "& pontos &"" + vbNewLine &_
	   "-------------------------")
if situcao = "Nivel dificil" then
	msgbox("-------------------------" + vbNewLine &_
	   "PARABENS!! "& nome &"" + vbNewLine &_
	   "Voce chegou no ultimo nivel"& nome &"" + vbNewLine &_
	   "-------------------------")
end if 
	   
call palavras
end sub

sub erro()
msgbox("-------------------------" + vbNewLine &_
	   "Nivel: "& situacao &"" + vbNewLine &_
	   "-------------------------" + vbNewLine &_
	   "Voce errou "& nome &" :(" + vbNewLine &_
	   "-------------------------" + vbNewLine &_
	   "Erros: "& qtde1 &"" + vbNewLine &_
	   "Acertos: "& qtde &"" + vbNewLine &_
	   "Pontos: "& pontos &"" + vbNewLine &_
	   "-------------------------")
call palavras
end Sub

#Include "TOTVS.ch"
#Include "FWMVCDef.ch"
#include "topconn.ch"

/*
Fonte....: zImpSG2
Descrição: Importação de dados para a tabela SG2 (Roteiro de Operações).
Autor....: Damião Carlos Duxe
Data.....: 27/05/2025
*/
User Function zImpSG2()

    Local oDlg
    Local oPanel
    
    Local oSay1, oSay2, oSay3, oSay4
    Local oGet1, oBtn1, oBtn2, oBtn3
    
    Local cArquivo := Space(200)
    Local nLinhas  := 0
   
    DEFINE MSDIALOG oDlg TITLE "Importação de Dados - Tabela SG2" FROM 000,000 TO 350,600 PIXEL
    
    // Panel principal
    oPanel := TPanel():New(0,0,'',oDlg,,.F.,.F.,,,0,0,.T.,.F.)
    oPanel:Align := CONTROL_ALIGN_ALLCLIENT
    
    // Componentes da tela
    @ 020,010 SAY oSay1 PROMPT "Arquivo CSV:" SIZE 060,010 OF oPanel PIXEL
    @ 018,070 MSGET oGet1 VAR cArquivo SIZE 150,012 OF oPanel PIXEL READONLY
    @ 018,225 BUTTON oBtn1 PROMPT "..." SIZE 015,012 OF oPanel PIXEL ACTION (cArquivo := fSelArquivo())
    
    @ 050,010 SAY oSay2 PROMPT "Selecione o arquivo CSV (separado por ponto e vírgula)" SIZE 200,010 OF oPanel PIXEL
    @ 065,010 SAY oSay3 PROMPT "para importação dos dados na tabela SG2 (Roteiro de Operações)." SIZE 200,010 OF oPanel PIXEL
    @ 080,010 SAY oSay4 PROMPT "Converta o Excel para CSV antes de importar." SIZE 200,010 OF oPanel PIXEL
    
    @ 120,050 BUTTON oBtn2 PROMPT "Importar" SIZE 060,020 OF oPanel PIXEL ACTION (nLinhas := fImportarDados(cArquivo), oDlg:End())
    @ 120,130 BUTTON oBtn3 PROMPT "Cancelar" SIZE 060,020 OF oPanel PIXEL ACTION oDlg:End()
    
    ACTIVATE MSDIALOG oDlg CENTERED
    
    If nLinhas > 0
        MsgInfo("Importação concluída com sucesso!" + CRLF + "Total de registros importados: " + cValToChar(nLinhas), "Sucesso")
    EndIf

Return

Static Function fSelArquivo()
    
    Local cArquivo := ""
    Local cTipos   := "Arquivos CSV (*.csv)|*.csv|Arquivos TXT (*.txt)|*.txt|"
    
    cArquivo := cGetFile(cTipos, "Selecione o arquivo CSV", 1, "", .T., GETF_LOCALHARD + GETF_NETWORKDRIVE)
    
Return cArquivo

Static Function fImportarDados(cArquivo)
    
    Local nLinhas := 0
    Local aDados  := {}
    Local lRet    := .F.
    
    If Empty(cArquivo)
        MsgAlert("Selecione um arquivo para importação!", "Atenção")
        Return 0
    EndIf
    
    If !File(cArquivo)
        MsgAlert("Arquivo não encontrado!", "Erro")
        Return 0
    EndIf
    
    // Processa o arquivo CSV
    Processa({|| lRet := fProcessaCsv(cArquivo, @aDados)}, "Processando arquivo CSV...")
    
    If !lRet .Or. Len(aDados) == 0
        MsgAlert("Erro ao processar o arquivo ou arquivo vazio!", "Erro")
        Return 0
    EndIf
    
    // Valida os dados
    Processa({|| lRet := fValidaDados(@aDados)}, "Validando dados...")
    
    If !lRet
        Return 0
    EndIf
    
    // Grava os dados na SG2
    Processa({|| nLinhas := fGravaDados(aDados)}, "Gravando dados na tabela SG2...")
    
Return nLinhas

Static Function fProcessaCsv(cArquivo, aDados)
    
    Local cBuffer      := ""
    Local cLinha       := ""
    Local cCampo       := ""
    Local cErros       := ""
    Local cSeparador   := ";"

    Local nHandle      := 0
    Local nTotalLinhas := 0
    Local nI           := 0
    Local nJ           := 0    
    Local nTamanho     := 0    

    Local aLinhas      := {}
    Local aLinha       := {}
    Local aCampos      := {}

    Local lRet         := .T.
   
    aDados := {}
    
    // Abre o arquivo
    nHandle := FOpen(cArquivo, 0)
    
    If nHandle == -1
        MsgAlert("Erro ao abrir o arquivo: " + cArquivo, "Erro")
        Return .F.
    EndIf
    
    // Lê todo o conteúdo do arquivo
    nTamanho := FSeek(nHandle, 0, 2)
    FSeek(nHandle, 0, 0)
    
    cBuffer := Space(nTamanho)
    FRead(nHandle, @cBuffer, nTamanho)
    FClose(nHandle)
    
    // Remove caracteres nulos
    cBuffer := StrTran(cBuffer, Chr(0), "")
    
    // Separa as linhas
    cBuffer := StrTran(cBuffer, Chr(13) + Chr(10), Chr(10))
    cBuffer := StrTran(cBuffer, Chr(13), Chr(10))
    
    aLinhas := StrTokArr(cBuffer, Chr(10))
    
    If Len(aLinhas) == 0
        MsgAlert("Arquivo vazio ou formato inválido!", "Erro")
        Return .F.
    EndIf
    
    // Remove a primeira linha (cabeçalho)
    If Len(aLinhas) > 1
        ADel(aLinhas, 1)
        ASize(aLinhas, Len(aLinhas) - 1)
    Else
        MsgAlert("Arquivo contém apenas o cabeçalho!", "Erro")
        Return .F.
    EndIf
    
    ProcRegua(Len(aLinhas))
    
    // Processa cada linha
    For nI := 1 To Len(aLinhas)
        
        IncProc("Processando linha: " + cValToChar(nI))
        
        cLinha := aLinhas[nI]
        
        // Remove apenas quebras de linha, mantém espaços
        cLinha := StrTran(cLinha, Chr(13), "")
        cLinha := StrTran(cLinha, Chr(10), "")
        
        // Pula linhas completamente vazias
        If Empty(AllTrim(cLinha))
            Loop
        EndIf
        
        // Separa os campos preservando campos vazios
        aCampos := fSeparaCampos(cLinha, ";")
        
        // Se não conseguiu separar adequadamente, tenta vírgula
        If Len(aCampos) < 5
            aCampos := fSeparaCampos(cLinha, ",")
            cSeparador := ","
        EndIf
        
        // Se ainda não conseguiu, tenta tab
        If Len(aCampos) < 5
            aCampos := fSeparaCampos(cLinha, Chr(9))
            cSeparador := "TAB"
        EndIf
        
        // Cria array com 32 posições
        aLinha := Array(32)
        AFill(aLinha, "")
        
        // Processa cada campo
        For nJ := 1 To Min(Len(aCampos), 32)
            
            // Limpa o campo
            cCampo := AllTrim(StrTran(aCampos[nJ], '"', ''))
            
            // Colunas 1 a 5: obrigatórias
            If nJ <= 5
                If Empty(cCampo)
                    cErros += "Linha " + cValToChar(nI+1) + ", Coluna " + cValToChar(nJ) + ": Campo obrigatório vazio" + CRLF
                EndIf
                aLinha[nJ] := cCampo
            Else
                // Colunas 6 a 32: opcionais
                If Empty(cCampo)
                    aLinha[nJ] := " "
                Else
                    aLinha[nJ] := cCampo
                EndIf
            EndIf
            
        Next nJ
        
        // Verifica se a primeira coluna tem dados válidos
        If Empty(AllTrim(aLinha[1]))
            Exit
        EndIf
        
        AAdd(aDados, aLinha)
        nTotalLinhas++
        
    Next nI
    
    // Se houver erros, exibe
    If !Empty(cErros)
        MsgAlert("Erros encontrados:" + CRLF + CRLF + cErros, "Erro de Validação")
        Return .F.
    EndIf
       
Return lRet

// Função para separar campos preservando campos vazios
Static Function fSeparaCampos(cLinha, cSeparador)
    
    Local cCampo  := ""

    Local nPos    := 0
    Local nInicio := 1

    Local aCampos := {}     
    
    // Adiciona separador no final para facilitar o processamento
    cLinha += cSeparador
    
    While .T.
        nPos := At(cSeparador, SubStr(cLinha, nInicio))
        
        If nPos == 0
            Exit
        EndIf
        
        // Extrai o campo (pode ser vazio)
        cCampo := SubStr(cLinha, nInicio, nPos - 1)
        AAdd(aCampos, cCampo)
        
        // Move para o próximo campo
        nInicio += nPos
        
    EndDo
    
Return aCampos

Static Function fValidaDados(aDados)

    Local cErros    := ""
    Local cProduto  := ""
    Local cRecurso  := ""
    Local cFil      := ""    
 
    Local nI        := 0
     
    Local lRet      := .T.
            
    ProcRegua(Len(aDados))
    
    For nI := 1 To Len(aDados)
        IncProc("Validando linha: " + cValToChar(nI))
        
        cFil     := AllTrim(cValToChar(aDados[nI][1])) // Coluna A - Filial
        cProduto := AllTrim(cValToChar(aDados[nI][3])) // Coluna C - Produto
        cRecurso := AllTrim(cValToChar(aDados[nI][5])) // Coluna E - Recurso
        
        // Validações obrigatórias
        If Empty(cProduto)
            cErros += "Linha " + cValToChar(nI+1) + ": Produto não informado" + CRLF
        EndIf
        
        If Empty(cRecurso)
            cErros += "Linha " + cValToChar(nI+1) + ": Recurso não informado" + CRLF
        EndIf
        
        // Valida se produto existe na SB1
        DbSelectArea("SB1")
        SB1->(DbSetOrder(1)) // B1_FILIAL + B1_COD
        If !Empty(cProduto) .And. !SB1->(DbSeek(xFilial("SB1") + cProduto))
            cErros += "Linha " + cValToChar(nI+1) + ": Produto " + cProduto + " não cadastrado" + CRLF
        EndIf
        
        // Valida se recurso existe na SH1
        DbSelectArea("SH1")
        SH1->(DbSetOrder(1)) // H1_FILIAL + H1_CODIGO
        If !Empty(cRecurso) .And. !SH1->(DbSeek(cFil + cRecurso))
            cErros += "Linha " + cValToChar(nI+1) + ": Recurso " + cRecurso + " não cadastrado" + CRLF
        EndIf
    Next nI
    
    If !Empty(cErros)
        lRet := .F.
        fMostraErros(cErros)
    EndIf
    
Return lRet

Static Function fMostraErros(cErros)

    Local oDlg, oMemo
    
    DEFINE MSDIALOG oDlg TITLE "Erros de Validação" FROM 000,000 TO 400,600 PIXEL
    
    @ 010,010 GET oMemo VAR cErros MEMO SIZE 280,160 OF oDlg PIXEL READONLY
    @ 180,250 BUTTON "Fechar" SIZE 040,015 OF oDlg PIXEL ACTION oDlg:End()
    
    ACTIVATE MSDIALOG oDlg CENTERED

Return

Static Function fGravaDados(aDados)

    Local cDescricao := ""
    
    Local nLinhas    := 0
    Local nI         := 0
        
    ProcRegua(Len(aDados))
    
    Begin Transaction
    
        For nI := 1 To Len(aDados)
            IncProc("Gravando linha: " + cValToChar(nI))
            
            // Busca descrição do recurso na SH1
            cDescricao := ""
            DbSelectArea("SH1")
            SH1->(DbSetOrder(1))
            If SH1->(DbSeek(AllTrim(cValToChar(aDados[nI][1])) + AllTrim(cValToChar(aDados[nI][5]))))
                cDescricao := SH1->H1_DESCRI
            EndIf
            
            // Grava registro na SG2
            DbSelectArea("SG2")
            SG2->(DbSetOrder(1)) // G2_FILIAL + G2_PRODUTO + G2_CODIGO + G2_TRT

            If Empty(Alltrim(GetAdvFVal( "SG2", "G2_PRODUTO", AllTrim(cValToChar(aDados[nI][1])) + ;
                                                              AllTrim(cValToChar(aDados[nI][3])) + ;
                                                              "01" + ;
                                                              AllTrim(cValToChar(aDados[nI][2])), 1, "")))
            
                RecLock("SG2", .T.)

                    SG2->G2_FILIAL  := AllTrim(cValToChar(aDados[nI][1]))          // Coluna A
                    SG2->G2_OPERAC  := AllTrim(cValToChar(aDados[nI][2]))          // Coluna B
                    SG2->G2_CODIGO  := "01"        
                    SG2->G2_PRODUTO := AllTrim(cValToChar(aDados[nI][3]))          // Coluna C 
                    SG2->G2_CTRAB   := AllTrim(cValToChar(aDados[nI][4]))          // Coluna D
                    SG2->G2_RECURSO := AllTrim(cValToChar(aDados[nI][5]))          // Coluna E
                    SG2->G2_FERRAM  := AllTrim(cValToChar(aDados[nI][6]))          // Coluna F
                    SG2->G2_LINHAPR := AllTrim(cValToChar(aDados[nI][7]))          // Coluna G
                    SG2->G2_TPLINHA := AllTrim(cValToChar(aDados[nI][8]))          // Coluna H
                    SG2->G2_DESCRI  := AllTrim(cDescricao)                         // Coluna I
                    SG2->G2__QTDPRO := Val(aDados[nI][10])                         // Coluna J
                    SG2->G2__TMPPRO := Val(aDados[nI][11])                         // Coluna K
                    SG2->G2__DTAFER := Date()                                      // Coluna L
                    SG2->G2__CUSPRO := Val(aDados[nI][13])                         // Coluna M
                    SG2->G2_MAOOBRA := Val(aDados[nI][14])                         // Coluna N
                    SG2->G2_SETUP   := Val(aDados[nI][15])                         // Coluna O
                    SG2->G2_TPOPER  := AllTrim(cValToChar(aDados[nI][16]))         // Coluna P
                    SG2->G2_TPDESD  := AllTrim(cValToChar(aDados[nI][17]))         // Coluna Q
                    SG2->G2_LOTEPAD := Val(aDados[nI][18])                         // Coluna R
                    SG2->G2_TEMPAD  := Val(aDados[nI][19])                         // Coluna S
                    SG2->G2_TEMPSOB := Val(aDados[nI][20])                         // Coluna T
                    SG2->G2_TEMPDES := Val(aDados[nI][21])                         // Coluna U
                    SG2->G2_DESPROP := AllTrim(cValToChar(aDados[nI][22]))         // Coluna V
                    SG2->G2_ROTALT  := AllTrim(cValToChar(aDados[nI][23]))         // Coluna W
                    SG2->G2_TPALOCF := AllTrim(cValToChar(aDados[nI][24]))         // Coluna X
                    SG2->G2_TEMPEND := Val(aDados[nI][25])                         // Coluna Y
                    SG2->G2_DEPTO   := AllTrim(cValToChar(aDados[nI][26]))         // Coluna Z
                    SG2->G2_DTINI   := sToD(aDados[nI][27])                        // Coluna AA
                    SG2->G2_DTFIM   := sToD(aDados[nI][28])                        // Coluna AB
                    SG2->G2_LISTA   := AllTrim(cValToChar(aDados[nI][29]))         // Coluna AC                                             
                    SG2->G2__USRAFE := AllTrim(RetCodUsr())                        // Coluna AD
                    SG2->G2_USAALT  := AllTrim(cValToChar(aDados[nI][31]))         // Coluna AE
                    SG2->G2_OPE_OBR := "S"
                    SG2->G2_SEQ_OBR := "S"
                    SG2->G2_LAU_OBR := "S"   
                    
                MsUnlock()
                
                nLinhas++
            Else
                FWAlertWarning("Registro já existe na SG2. " + Chr(10) + ;
                "Filial..: " + AllTrim(cValToChar(aDados[nI][1])) + Chr(10) + ;
                "Produto.: " + AllTrim(cValToChar(aDados[nI][3])) + Chr(10) + ;
                "Código..: 01" + Chr(10) + ;
                "Operação: " + AllTrim(cValToChar(aDados[nI][2])), "zImpSG2"  )

            Endif
        Next nI
    
    End Transaction
    
Return nLinhas



{utp/utapi019.i}

DEFINE VARIABLE c-anexo      AS CHARACTER FORMAT "x(50)".
DEFINE VARIABLE c-horas      AS CHARACTER   NO-UNDO.
DEFINE VARIABLE i-hora-atual AS INTEGER     NO-UNDO.

IF NOT AVAIL param_email THEN
    FIND FIRST param_email NO-LOCK NO-ERROR.

    
FOR EACH tt-envio2:   DELETE tt-envio2.   END.
FOR EACH tt-mensagem: DELETE tt-mensagem. END.
FOR EACH tt-erros:    DELETE tt-erros.    END.


ASSIGN c-horas = STRING(TIME,"HH:MM:SS")
       i-hora-atual = INT(ENTRY(1,c-horas,":")).

ASSIGN c-anexo = "c:\tempo\anexo.xlsx".

CREATE tt-envio2.
ASSIGN tt-envio2.versao-integracao = 1
       tt-envio2.servidor          = param_email.cod_servid_e_mail
       tt-envio2.porta             = param_email.num_porta
       tt-envio2.exchange          = NO
       tt-envio2.remetente         = "remetente@progress.com.br"
       tt-envio2.destino           = "destino@progress.com.br"
       tt-envio2.assunto           = "Envio de Email"
       tt-envio2.importancia       = 2
       tt-envio2.formato           = "texto".
       tt-envio2.arq-anexo         = c-anexo.

IF i-hora-atual <= 12 THEN DO:
    RUN pi-cria-mensagem(INPUT "Bom dia." + CHR(13) + CHR(10)).
END.
ELSE DO:
    RUN pi-cria-mensagem(INPUT "Boa tarde." + CHR(13) + CHR(10)).
END.

RUN pi-cria-mensagem(INPUT "Segue relatório." + CHR(13) + CHR(10) + CHR(10)).
RUN pi-cria-mensagem(INPUT "Atenciosamente," + CHR(13) + CHR(10)).

RUN utp/utapi019.p PERSISTENT SET h-utapi019.
RUN pi-execute2 IN h-utapi019 (INPUT  TABLE tt-envio2,
                               INPUT  TABLE tt-mensagem,
                               OUTPUT TABLE tt-erros).

DELETE PROCEDURE h-utapi019.   

PROCEDURE pi-cria-mensagem :

    DEF INPUT PARAMETER cMensagem AS CHAR NO-UNDO.
    DEFINE VARIABLE i-linha     AS INTEGER     NO-UNDO.

    ASSIGN i-linha = i-linha + 1.

    CREATE tt-mensagem.
    ASSIGN tt-mensagem.seq-mensagem = i-linha
           tt-mensagem.mensagem     = cMensagem.

END PROCEDURE.

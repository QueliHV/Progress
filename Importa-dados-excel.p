
DEFINE TEMP-TABLE tt-dados NO-UNDO
    FIELD it-codigo  LIKE ITEM.it-codigo
    FIELD descricao  LIKE ITEM.desc-item
    FIELD un         LIKE ITEM.un
    FIELD peso-bruto LIKE ITEM.peso-bruto.

DEFINE VARIABLE c-arquivo-imp      AS CHARACTER FORMAT "x(50)" NO-UNDO.
DEFINE VARIABLE l-arquivo-ok       AS LOGICAL                  NO-UNDO.
DEFINE VARIABLE h-acomp            AS HANDLE                   NO-UNDO.
DEFINE VARIABLE i-lin              AS INTEGER                  NO-UNDO.
DEFINE VARIABLE chExcelApplication AS COM-HANDLE               NO-UNDO.
DEFINE VARIABLE chWorkbook         AS COM-HANDLE               NO-UNDO.
DEFINE VARIABLE chWorksheet        AS COM-HANDLE               NO-UNDO.

ASSIGN c-arquivo-imp = "arquivo.xls".

SYSTEM-DIALOG GET-FILE c-arquivo-imp 
     TITLE "Arquivo a Importar" 
     FILTERS "arquivo xls (*.xls)" "*.xls", 
             "arquivo xlsx (*.xlsx)" "*.xlsx",
             "todos arquivos (*.*)"   "*.*"
     INITIAL-DIR SESSION:TEMP-DIRECTORY
     MUST-EXIST 
     USE-FILENAME 
     DEFAULT-EXTENSION "xls"
     UPDATE l-arquivo-ok. 

EMPTY TEMP-TABLE tt-dados.

IF NOT l-arquivo-ok THEN
    RETURN NO-APPLY.

DEF VAR c-linha AS CHAR NO-UNDO.

IF NOT VALID-HANDLE(h-acomp) THEN
    RUN utp/ut-acomp.p PERSISTENT SET h-acomp.

RUN pi-inicializar IN h-acomp ("Importando Dados").

ASSIGN i-lin = 1.       

DO ON STOP UNDO, LEAVE:

    CREATE "Excel.Application" chExcelApplication NO-ERROR.
    ASSIGN chExcelApplication:VISIBLE = FALSE
           chWorkbook                 = chExcelApplication:WorkBooks:OPEN(SEARCH(c-arquivo-imp))
           chWorksheet                = chWorkbook:Sheets(1).

    blk_imp:
    REPEAT:
        ASSIGN i-lin = i-lin + 1.       

        IF  chWorksheet:range("B" + STRING(i-lin)):VALUE = "" OR 
            chWorksheet:range("B" + STRING(i-lin)):VALUE = ?  THEN
            LEAVE blk_imp.

        CREATE tt-dados.
        ASSIGN tt-dados.it-codigo  = STRING(chWorksheet:range("A" + STRING(i-lin)):TEXT)
               tt-dados.descricao  = STRING(chWorksheet:range("B" + STRING(i-lin)):TEXT)
               tt-dados.un         = STRING(chWorksheet:range("C" + STRING(i-lin)):TEXT)
               tt-dados.peso-bruto = DEC(STRING(chWorksheet:range("D" + STRING(i-lin)):TEXT)).

        RUN pi-acompanhar IN h-acomp ("Importando itens: " + STRING(tt-dados.it-codigo) + " - " + STRING(tt-dados.descricao)).
    END.

    RELEASE OBJECT chWorksheet.
    RELEASE OBJECT chWorkbook.
    chExcelApplication:QUIT().
    RELEASE OBJECT chExcelApplication.

END. /* DO ON STOP UNDO, LEAVE */


RUN pi-finalizar IN h-acomp.

// Lista dados importador
/*FOR EACH tt-dados:

    DISP tt-dados.it-codigo 
         tt-dados.descricao 
         tt-dados.un        
         tt-dados.peso-bruto.
END.*/
                             

MESSAGE "Importa‡Æo Conclu¡da"
  VIEW-AS ALERT-BOX INFO BUTTONS OK.








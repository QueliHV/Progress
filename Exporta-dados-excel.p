DEFINE VARIABLE chExcel         AS COMPONENT-HANDLE          NO-UNDO.
DEFINE VARIABLE i-lin           AS INTEGER                   NO-UNDO.
DEFINE VARIABLE c-arquivo-saida AS CHARACTER FORMAT "x(50)"  NO-UNDO.
DEFINE VARIABLE h-acomp         AS HANDLE      NO-UNDO.

IF NOT VALID-HANDLE(h-acomp) THEN
    RUN utp/ut-acomp.p PERSISTENT SET h-acomp.

RUN pi-inicializar IN h-acomp ("Gerando Excel").

CREATE 'Excel.Application' chExcel.
IF chExcel = ? THEN LEAVE.
chExcel:VISIBLE=FALSE.
chExcel:Workbooks:ADD.


// Cabe‡alho Relat¢rio
ASSIGN i-lin = 1.
chExcel:Range("A" + STRING(i-lin)):VALUE = "C¢d Item".
chExcel:RANGE('A:A'):NUMBERFORMAT = '@'.
chExcel:RANGE('A:A'):ColumnWidth = 14.
chExcel:Range("B" + STRING(i-lin)):VALUE = "Descri‡Æo".
chExcel:RANGE('B:B'):ColumnWidth = 69.
chExcel:Range("C" + STRING(i-lin)):VALUE = "Un".
chExcel:RANGE('C:C'):ColumnWidth = 5.
chExcel:Range("D" + STRING(i-lin)):VALUE = "Peso Bruto".
chExcel:RANGE('D:D'):ColumnWidth = 14.
chExcel:Range("A1:D1"):SELECT.
chExcel:SELECTION:FONT:bold = TRUE.

FOR EACH ITEM:

    RUN pi-acompanhar IN h-acomp ("Lendo Itens: " + STRING(ITEM.it-codigo)).

    ASSIGN i-lin = i-lin + 1.
    chExcel:Range("A" + STRING(i-lin)):VALUE = ITEM.it-codigo.
    chExcel:Range("B" + STRING(i-lin)):VALUE = ITEM.desc-item.
    chExcel:Range("C" + STRING(i-lin)):VALUE = ITEM.un. 
    chExcel:Range("D" + STRING(i-lin)):VALUE = ITEM.peso-bruto.   

END.


// Verifica versÆo Excel
IF chExcel:VERSION >= "12" THEN        
    ASSIGN c-arquivo-saida = session:TEMP-DIRECTORY + "arquivo-" + STRING(TIME) + ".xlsx".
ELSE
    ASSIGN c-arquivo-saida = session:TEMP-DIRECTORY + "arquivo-" + STRING(TIME) + ".xls".


chExcel:ActiveWorkbook:SaveAs(c-arquivo-saida,56,"","",FALSE,FALSE,3,1,1,1).
ASSIGN chExcel:ScreenUpdating = YES
       chExcel:VISIBLE        = YES.

RUN pi-finalizar IN h-acomp.

// Elimina handle da memoria
IF VALID-HANDLE(chExcel) THEN
    RELEASE OBJECT chExcel.

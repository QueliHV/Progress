{utp/ut-api.i}
{utp/ut-api-action.i REST_GET_todos    GET /ALL* }
{utp/ut-api-action.i REST_GET_item GET /~* }
{utp/ut-api-notfound.i}
{method/dbotterr.i}

DEFINE TEMP-TABLE Erro NO-UNDO LIKE RowErrors.      

PROCEDURE REST_GET_item:

    DEF INPUT PARAM jsonInput        AS JsonObject NO-UNDO.
    DEF OUTPUT PARAM jsonOutput      AS JsonObject NO-UNDO.  
    DEFINE VARIABLE jsonArrayErroOut AS JsonArray  NO-UNDO.
    DEFINE VARIABLE jsonErro         AS JsonObject NO-UNDO.

    DEFINE VARIABLE aPathParams        AS JsonArray            NO-UNDO.
    DEFINE VARIABLE oRequestParser     AS JsonAPIRequestParser NO-UNDO.
    DEFINE VARIABLE cJsonArrayChar     AS CHARACTER            NO-UNDO.

    DEFINE VARIABLE c-cod-item AS CHARACTER   NO-UNDO.

    /** Importa o json recebido **/
    oRequestParser = NEW JsonAPIRequestParser(jsonInput).

    /** Recebe um array com os parƒmetros enviados na chamada **/
    aPathParams = oRequestParser:getPathParams().    

    /** Copia todo o conteudo do array para um campo char separado por "," **/
    ASSIGN cJsonArrayChar = JsonAPIUtils:getJsonArrayChar(aPathParams).
    ASSIGN c-cod-item = ENTRY(01,cJsonArrayChar,",").

    FIND FIRST ITEM NO-LOCK
         WHERE ITEM.it-codigo = c-cod-item NO-ERROR.
    IF AVAIL ITEM THEN DO:

        jsonOutput = NEW JsonObject().
        jsonOutput:ADD("C¢digo", ITEM.it-codigo).
        jsonOutput:ADD("Descri‡Æo", ITEM.desc-item).

    END.
    ELSE DO:

        CREATE Erro.
        ASSIGN Erro.ErrorNumber = 010
               Erro.ErrorSubType = "ERROR"
               Erro.ErrorType    = "EMS"
               Erro.ErrorHelp    = "Item inv lido"
               Erro.ErrorDescription = "Item nÆo cadastrado!".

    END.

    jsonArrayErroOut = NEW JsonArray().

    FOR EACH Erro:
        jsonErro = NEW JsonObject().
        jsonErro:ADD("ErrorSubType", Erro.ErrorNumber).
        jsonErro:ADD("ErrorType", Erro.ErrorType).
        jsonErro:ADD("ErrorDescription", Erro.ErrorDescription).
        jsonErro:ADD("ErrorNumber", STRING(Erro.ErrorNumber)).
        jsonErro:ADD("ErrorHelp", STRING(Erro.ErrorHelp)).
        jsonArrayErroOut:ADD(jsonErro).
    END.
    
    // Json de retorno        
    jsonOutput:Add("Erro", jsonArrayErroOut).

END PROCEDURE.




// AAToken: Returns the Control Room JWT token (text)
(AACR as text, AAUser as text, AAApiKey as text) as text =>
let
    Url   = AACR & "/v1/authentication",
    Body  = [ username = AAUser, apiKey = AAApiKey ],
    Resp  = Json.Document(
              Web.Contents(
                Url,
                [
                  Headers = [#"Content-Type"="application/json"],
                  Content = Text.ToBinary(Json.FromValue(Body))
                ]
              )
            ),
    Token = Resp[token]
in
    Token



----------------------------------


// AAScheduledList: Returns a table of scheduled automations
(AACR as text, Token as text) as table =>
let
    Url     = AACR & "/v1/schedule/automations/list",
    Raw     = Json.Document(
                Web.Contents(
                  Url,
                  [
                    Headers = [
                      #"Content-Type"="application/json",
                      #"X-Authorization"= Token
                    ],
                    Content = Text.ToBinary("{}")
                  ]
                )
              ),
    // Many AA endpoints return { "list": [ {...}, {...} ], "total": n, ... }
    Items   = if Record.HasFields(Raw, "list") then Raw[list] else Raw,
    Tbl     = Table.FromList(Items, Splitter.SplitByNothing(), {"Record"}, null, ExtraValues.Error),
    Expanded= try
                Table.ExpandRecordColumn(
                  Tbl, "Record",
                  {"id","name","description","status","nextRunTime","scheduleType","cronExpression","runAsUser","devicePool","createdBy","createdOn","updatedOn"},
                  {"id","name","description","status","nextRunTime","scheduleType","cronExpression","runAsUser","devicePool","createdBy","createdOn","updatedOn"}
                )
              otherwise
                // Fallback: if fields differ, let you expand dynamically in Power BI
                Tbl
in
    Expanded



-----------

// AAToken (fixed): Returns the Control Room JWT token (text)
(AACR as text, AAUser as text, AAApiKey as text) as text =>
let
    Url   = AACR & "/v1/authentication",
    Body  = [ username = AAUser, apiKey = AAApiKey ],
    Resp  = Json.Document(
              Web.Contents(
                Url,
                [
                  Headers = [#"Content-Type"="application/json"],
                  Content = Json.FromValue(Body)   // <-- no Text.ToBinary
                ]
              )
            ),
    // Some CRs return {"token":"..."}; others might return a raw string.
    Token =
        if (Resp is record) and Record.HasFields(Resp, "token")
        then Text.From(Resp[token])
        else
            // fallback if the endpoint returns raw text instead of JSON
            let raw = Web.Contents(Url, [Headers=[#"Content-Type"="application/json"], Content=Json.FromValue(Body)])
            in Text.FromBinary(raw, TextEncoding.Utf8)
in
    Token


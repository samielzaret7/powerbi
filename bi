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


-------------------------

// AAUsersList (simple version)
(AACR as text, Token as text) as table =>
let
    Url  = AACR & "/v1/users/list",
    Raw  = Json.Document(
              Web.Contents(
                  Url,
                  [
                      Headers = [
                          #"Content-Type"    = "application/json",
                          #"X-Authorization" = Token
                      ],
                      Content = Json.FromValue([ page = 0, size = 100 ])
                  ]
              )
           ),
    // Some Control Rooms return { list=[...], total=n }
    Items = if Raw is record and Record.HasFields(Raw, "list") then Raw[list] else Raw,
    Tbl   = Table.FromList(Items, Splitter.SplitByNothing(), {"Record"})
in
    Tbl


--------------------


// AAUsersList: returns a table of all users from /v1/users/list
// Params:
//   AACR  - Control Room base URL, e.g. "https://cr.company.com"
//   Token - X-Authorization token from AAToken()
//   PageSize (optional) - defaults to 100
(AACR as text, Token as text, optional PageSize as nullable number) as table =>
let
    Size   = if PageSize <> null then PageSize else 100,
    Url    = AACR & "/v1/users/list",

    FetchPage = (p as number) as record =>
        let
            Body    = [ page = p, size = Size ],
            Raw     = Json.Document(
                        Web.Contents(
                            Url,
                            [
                                Headers = [
                                    #"Content-Type"    = "application/json",
                                    #"X-Authorization" = Token
                                ],
                                Content = Json.FromValue(Body)
                            ]
                        )
                      ),
            // Common shapes: { list=[...], total=n } OR just [ ... ]
            ListOut = if Raw is record and Record.HasFields(Raw, "list") then Raw[list] else Raw,
            Total   = if Raw is record and Record.HasFields(Raw, "total") then Raw[total] else null,
            Count   = if ListOut is list then List.Count(ListOut) else 0
        in
            [items = ListOut, count = Count, total = Total],

    Pages =
        List.Generate(
            () => [p = 0, r = FetchPage(0)],
            each [r][count] > 0,
            each [p = [p] + 1, r = FetchPage([p] + 1)],
            each [r][items]
        ),

    AllItems = List.Combine(Pages),

    // Turn into a table and expand common fields. Adjust if your CR has extra fields.
    Tbl0 = Table.FromList(AllItems, Splitter.SplitByNothing(), {"Record"}, null, ExtraValues.Error),
    Tbl  = Table.ExpandRecordColumn(
              Tbl0, "Record",
              {"id","username","email","firstName","lastName","isActive","roles","createdOn","updatedOn"},
              {"id","username","email","firstName","lastName","isActive","roles","createdOn","updatedOn"}
          ),
    // Optional: flatten roles into a comma-separated list of role names
    RolesExpanded =
        if Table.HasColumns(Tbl, "roles")
        then
            let
                AddRoleNames = Table.AddColumn(
                    Tbl, "roleNames",
                    each if [roles] is list
                         then Text.Combine(List.Transform([roles], (r) => try r[name] otherwise Text.From(r)), ", ")
                         else null,
                    type text
                ),
                DropRolesCol = Table.RemoveColumns(AddRoleNames, {"roles"})
            in
                DropRolesCol
        else
            Tbl
in
    RolesExpanded


-----------------

// AAUsersList_Simple: minimal loader for /v1/usermanagement/users/list
// Usage:
//   let
//     Token = AAToken(AACR, AAUser, AAApiKey),
//     Users = AAUsersList_Simple(AACR, Token)   // default body: [page=0, size=100]
//   in Users
(AACR as text, Token as text, optional Body as nullable record) as table =>
let
    Url     = AACR & "/v1/usermanagement/users/list",
    Payload = if Body <> null then Body else [ page = 0, size = 100 ],
    Source  = Web.Contents(
                Url,
                [
                    Headers = [
                        #"Content-Type"    = "application/json",
                        #"X-Authorization" = Token
                    ],
                    Content = Json.FromValue(Payload)
                ]
              ),
    ParsedTry = try Json.Document(Source) otherwise null,
    Parsed    = if ParsedTry <> null then ParsedTry else error "Users API: response was not JSON.",

    // If the API replied with an error object, raise a readable error
    _ = if (Parsed is record) and Record.HasFields(Parsed, {"code","message"})
        then error Error.Record("Users API", Parsed[message], Parsed)
        else null,

    // Normalize common shapes:
    // 1) { list=[...], total=n }  2) { users=[...] }  3) [ ... ]  4) {} (empty)
    Items =
        if (Parsed is record) and Record.HasFields(Parsed, "list") then Parsed[list]
        else if (Parsed is record) and Record.HasFields(Parsed, "users") then Parsed[users]
        else if (Parsed is list) then Parsed
        else {},

    Tbl = Table.FromList(Items, Splitter.SplitByNothing(), {"Record"}, null, ExtraValues.Error)
in
    Tbl


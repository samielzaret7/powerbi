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



---------------------


// AABotExecutionsLastNDays: fetch executions from the last N days (default 90)
// Params:
//   AACR   : base URL, e.g. "https://your-cr.company.com"
//   Token  : X-Authorization from AAToken()
//   Days   : optional number of days back (defaults to 90)
(AACR as text, Token as text, optional Days as nullable number) as table =>
let
    NDays      = if Days <> null then Days else 90,

    // --- Date range (UTC) ---
    NowUtc     = DateTimeZone.UtcNow(),
    FromUtc    = DateTimeZone.ToUtc(DateTimeZone.FixedUtcNow()) - #duration(NDays,0,0,0),

    // Many AA v3 endpoints expect epoch millis in filters:
    ToMs       = Number.RoundDown(1000 * Duration.TotalSeconds(NowUtc - #datetime(1970,1,1,0,0,0))),
    FromMs     = Number.RoundDown(1000 * Duration.TotalSeconds(FromUtc - #datetime(1970,1,1,0,0,0))),

    Url        = AACR & "/v3/activity/list",

    // Page settings
    PageLen    = 200,

    // Build request body using the common v3 filter schema
    MakeBody = (offset as number) as record =>
        [
          filter = [
            operator = "AND",
            operands = {
              [ field = "createdOn", operator = "GREATER_THAN_EQUALS", value = FromMs ],
              [ field = "createdOn", operator = "LESS_THAN_EQUALS",  value = ToMs   ]
            }
          ],
          sort   = { [ field = "createdOn", direction = "desc" ] },
          page   = [ offset = offset, length = PageLen ]
        ],

    // Fetch one page
    FetchPage = (offset as number) as list =>
        let
            Body   = MakeBody(offset),
            Resp   = Json.Document(
                        Web.Contents(
                          Url,
                          [
                            Headers = [
                              #"Content-Type"    = "application/json",
                              #"Accept"          = "application/json",
                              #"X-Authorization" = Token
                            ],
                            Content = Json.FromValue(Body)
                          ]
                        )
                      ),
            Items =
              if (Resp is record) and Record.HasFields(Resp, "list") then Resp[list]
              else if (Resp is list) then Resp
              else {}
        in
            Items,

    // Iterate pages until a short page is returned
    Pages = List.Generate(
              () => [offset = 0, batch = FetchPage(0)],
              each List.Count([batch]) > 0,
              each [ offset = [offset] + PageLen, batch = FetchPage([offset] + PageLen) ],
              each [batch]
            ),

    AllItems = List.Combine(Pages),
    Tbl0     = Table.FromList(AllItems, Splitter.SplitByNothing(), {"Record"}),
    // Expand *all* fields dynamically
    Expanded =
        if Table.RowCount(Tbl0) > 0 then
            Table.ExpandRecordColumn(Tbl0, "Record", Record.FieldNames(Tbl0{0}[Record]))
        else
            Tbl0
in
    Expanded



------------------

// AABotExecutionsLastNDays (fixed datetime math)
(AACR as text, Token as text, optional Days as nullable number) as table =>
let
    NDays   = if Days <> null then Days else 90,

    // UTC times
    NowUtc  = DateTimeZone.UtcNow(),
    FromUtc = DateTimeZone.AddDays(NowUtc, -NDays),

    // Convert to epoch milliseconds
    EpochBase = #datetimezone(1970,1,1,0,0,0,+00:00),
    ToMs   = Number.RoundDown(1000 * Duration.TotalSeconds(NowUtc - EpochBase)),
    FromMs = Number.RoundDown(1000 * Duration.TotalSeconds(FromUtc - EpochBase)),

    Url     = AACR & "/v3/activity/list",
    PageLen = 200,

    MakeBody = (offset as number) as record =>
        [
          filter = [
            operator = "AND",
            operands = {
              [ field = "createdOn", operator = "GREATER_THAN_EQUALS", value = FromMs ],
              [ field = "createdOn", operator = "LESS_THAN_EQUALS",  value = ToMs ]
            }
          ],
          sort = { [ field = "createdOn", direction = "desc" ] },
          page = [ offset = offset, length = PageLen ]
        ],

    FetchPage = (offset as number) as list =>
        let
            Body  = MakeBody(offset),
            Resp  = Json.Document(
                      Web.Contents(
                        Url,
                        [
                          Headers = [
                            #"Content-Type"    = "application/json",
                            #"Accept"          = "application/json",
                            #"X-Authorization" = Token
                          ],
                          Content = Json.FromValue(Body)
                        ]
                      )
                    ),
            Items =
                if (Resp is record) and Record.HasFields(Resp, "list") then Resp[list]
                else if (Resp is list) then Resp
                else {}
        in
            Items,

    Pages = List.Generate(
              () => [offset = 0, batch = FetchPage(0)],
              each List.Count([batch]) > 0,
              each [ offset = [offset] + PageLen, batch = FetchPage([offset] + PageLen) ],
              each [batch]
            ),

    AllItems = List.Combine(Pages),
    Tbl0     = Table.FromList(AllItems, Splitter.SplitByNothing(), {"Record"}),

    // Dynamically expand all fields
    Expanded =
        if Table.RowCount(Tbl0) > 0 then
            Table.ExpandRecordColumn(Tbl0, "Record", Record.FieldNames(Tbl0{0}[Record]))
        else
            Tbl0
in
    Expanded



-------------

// AABotExecutionsLastNDays_ByEndTime
// Gets executions from the last N days using /v3/activity/list
// Filters: status = COMPLETED (change if you want), endDateTime BETWEEN [From, To]
// Sorts by endDateTime desc. Pages until empty/short page.
(AACR as text, Token as text, optional Days as nullable number, optional PageLen as nullable number, optional Status as nullable text) as table =>
let
    NDays      = if Days <> null then Days else 90,
    PageSize   = if PageLen <> null then PageLen else 1000,
    WantStatus = if Status <> null then Status else "COMPLETED",

    // UTC window
    NowUtc     = DateTimeZone.UtcNow(),
    FromUtc    = NowUtc - #duration(NDays, 0, 0, 0),

    // epoch ms
    EpochBase  = #datetimezone(1970,1,1,0,0,0,0,0),
    ToMs       = Number.RoundDown(1000 * Duration.TotalSeconds(NowUtc  - EpochBase)),
    FromMs     = Number.RoundDown(1000 * Duration.TotalSeconds(FromUtc - EpochBase)),

    Url        = AACR & "/v3/activity/list",

    // Build request body following common v3 schema
    MakeBody = (offset as number) as record =>
        [
          filter = [
            operator = "AND",
            operands = {
              [ field = "status",       operator = "EQ",                    value = WantStatus ],
              [ field = "endDateTime",  operator = "GREATER_THAN_EQUALS",   value = FromMs     ],
              [ field = "endDateTime",  operator = "LESS_THAN_EQUALS",      value = ToMs       ]
            }
          ],
          sort = { [ field = "endDateTime", direction = "desc" ] },
          page = [ offset = offset, length = PageSize ]
        ],

    FetchPage = (offset as number) as list =>
        let
            Body  = MakeBody(offset),
            Resp  = Json.Document(
                        Web.Contents(
                          Url,
                          [
                            Headers = [
                              #"Content-Type"    = "application/json",
                              #"Accept"          = "application/json",
                              #"X-Authorization" = Token
                            ],
                            Content = Json.FromValue(Body)
                          ]
                        )
                    ),
            Items =
                if (Resp is record) and Record.HasFields(Resp, "list") then Resp[list]
                else if (Resp is list) then Resp
                else {}
        in
            Items,

    Pages = List.Generate(
              () => [offset = 0, batch = FetchPage(0)],
              each List.Count([batch]) > 0,
              each 
                let nextOffset = [offset] + PageSize
                in  [offset = nextOffset, batch = FetchPage(nextOffset)],
              each [batch]
            ),

    AllItems = List.Combine(Pages),
    Tbl0     = Table.FromList(AllItems, Splitter.SplitByNothing(), {"Record"}),

    // Expand everything first so you can prune later
    Expanded =
        if Table.RowCount(Tbl0) > 0 then
            Table.ExpandRecordColumn(Tbl0, "Record", Record.FieldNames(Tbl0{0}[Record]))
        else
            Tbl0
in
    Expanded




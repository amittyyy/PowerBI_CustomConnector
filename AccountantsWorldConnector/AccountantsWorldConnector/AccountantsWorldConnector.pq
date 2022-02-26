/***************************************************************************************************************
**** AT. 020252022 Accoutnig power Customer Power BI Connector. ************************************************
**** This application will get data form Accoutning power External API via OAuth2 service. ********************* 
***************************************************************************************************************/

section AccountantsWorldConnector;

url = "https://dev-accountingexternalapi.accountingpower.com/api/";

//AT. 02252022 Ext Accounting power api client ID and Secrete
client_id = "36VXYCeMwpB0pHrmC6PZqkMdwu%2BYCCU%2FABmui1i60GjNOZ1oafQK9Zm5eKbOImoP";
client_secret = "LY5q9u8CjNUQQp7t69N%2BLQvpGtxj69JTLEcIAvTOyEasnAF3GYWYS8aAaW6ElsRl";

redirect_uri = "https://oauth.powerbi.com/views/oauthredirect.html";
token_uri = "https://dev-auth.accountantsoffice.com/connect/token";
authorize_uri = "https://dev-auth.accountantsoffice.com/connect/authorize";
logout_uri = "https://dev-auth.accountantsoffice.com/connect/logout";
// apiKey = "23911c167cc4417d940b4bb8f17c0a6d";

//AT. ACCounting client Api Keys
apiKey = "3911864bb1dd427a8983f2b10522f827";

// Login modal window dimensions
windowWidth = 800;
windowHeight = 1024;

[DataSource.Kind="AccountantsWorldConnector", Publish="AccountantsWorldConnector.Publish"]
shared AccountantsWorldConnector.Contents = Value.ReplaceType(AccountingTables, type function () as any);

// Data Source Kind description
AccountantsWorldConnector = [
 TestConnection = (dataSourcePath) => { "AccountantsWorldConnector.Contents", dataSourcePath },
    Authentication = [
        OAuth = [
            StartLogin=StartLogin,
            FinishLogin=FinishLogin,
            Refresh=Refresh,
            Logout=Logout            
        ]
    ],
    Label = Extension.LoadString("DataSourceLabel")
];

// Data Source UI publishing description
AccountantsWorldConnector.Publish = [
    Beta = true,
    Category = "Other",
    ButtonText = { Extension.LoadString("ButtonTitle"), Extension.LoadString("ButtonHelp") },
    LearnMoreUrl = "https://www.accountantsworld.com/",
    SourceImage = AccountantsWorldConnector.Icons,
    SourceTypeImage = AccountantsWorldConnector.Icons
];

// Helper functions for OAuth2: StartLogin, FinishLogin, Refresh, Logout
StartLogin = (resourceUrl, state, display) =>
    let
        authorizeUrl = authorize_uri & "?" & Uri.BuildQueryString([
            response_type = "code",
            client_id = client_id,  
            redirect_uri = redirect_uri,
            state = state,
            scope = "accounting_api"   //payroll_api
            
        ])
    in
        [
            LoginUri = authorizeUrl,
            CallbackUri = redirect_uri,
            WindowHeight = 720,
            WindowWidth = 1024,
            Context = null
        ];

FinishLogin = (context, callbackUri, state) =>
    let
        // parse the full callbackUri, and extract the Query string
        parts = Uri.Parts(callbackUri)[Query],
        // if the query string contains an "error" field, raise an error
        // otherwise call TokenMethod to exchange our code for an access_token
        result = if (Record.HasFields(parts, {"error", "error_description"})) then 
                    error Error.Record(parts[error], parts[error_description], parts)
                 else
                    TokenMethod("authorization_code", "code", parts[code])
    in
        result;

Refresh = (resourceUrl, refresh_token) => TokenMethod("refresh_token", "refresh_token", refresh_token);

Logout = (token) => logout_uri;

// see "Exchange code for access token: POST /oauth/token" at https://cloud.ouraring.com/docs/authentication for details
TokenMethod = (grantType, tokenField, code) =>
    let
        queryString = [
            grant_type = "authorization_code",
            redirect_uri = redirect_uri,
            client_id = client_id,
            client_secret = client_secret            
        ],
        queryWithCode = Record.AddField(queryString, tokenField, code),

        tokenResponse = Web.Contents(token_uri, [
            Content = Text.ToBinary(Uri.BuildQueryString(queryWithCode)),
            Headers = [
                #"Content-type" = "application/x-www-form-urlencoded",
                #"Accept" = "application/json"
            ],
            ManualStatusHandling = {400} 
        ]),
        body = Json.Document(tokenResponse),
        result = if (Record.HasFields(body, {"error", "error_description"})) then 
                    error Error.Record(body[error], body[error_description], body)
                 else
                    body
    in
        result;

Value.IfNull = (a, b) => if a <> null then a else b;

AccountantsWorldConnector.Icons = [
    Icon16 = { Extension.Contents("AccountantsWorldConnector16.png"), Extension.Contents("AccountantsWorldConnector20.png"), Extension.Contents("AccountantsWorldConnector24.png"), Extension.Contents("AccountantsWorldConnector32.png") },
    Icon32 = { Extension.Contents("AccountantsWorldConnector32.png"), Extension.Contents("AccountantsWorldConnector40.png"), Extension.Contents("AccountantsWorldConnector48.png"), Extension.Contents("AccountantsWorldConnector64.png") }
];

/***************************************************Api Call Functions starts from here **************************************************/   

GetCustomers = () as table =>
        let 

        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                RelativePath = "customers",
                Headers = headers
            ]
        ))[customers],
      data = Table.FromList(dataList, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
      customerColumns = Table.ExpandRecordColumn(data, "Column1", {"id", "name", "address1", "city", "state", "zip", "email", "phoneNumber", "balanceDue", "shipAddress1","shipCity", "shipState","shipZip","isInactive","salesTaxState", "salesTaxCounty", "salesTaxRate"})

    in
        customerColumns;

GetTrailBalance = () as table =>
        let 

        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                RelativePath = "TrialBalance",
                Query = [period="2021-04-30", IsYearToDate="true"],
                Headers = headers
            ]
        ))[accounts],
      data = Table.FromList(dataList, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
      trialBalanceColumns = Table.ExpandRecordColumn(data, "Column1",{"code","type","description","beginningBalance","transactionsDebit","transactionsCredit","unadjustedBalance","adjustmentsDebit","adjustmentsCredit","adjustedBalance","notes"})

    in
        trialBalanceColumns;


GetVendors = () as table =>
        let 

        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                RelativePath = "vendors",                
                Headers = headers
            ]
        ))[vendors],
      data = Table.FromList(dataList, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
      vendorsColumns = Table.ExpandRecordColumn(data, "Column1",{"id","name","aliasName","is1099","address1","city","phoneNumber","faxNumber","isInactive","terms","defaultAccountCode","departmentId","isW9"})

    in
        vendorsColumns;


//AT. 02252022 Combining all api return table into one table.
AccountingTables = () as table =>
    let
        source = #table({"Name", "Data"}, {
            { "Customers", GetCustomers() },
            { "TrailBalances", GetTrailBalance() },
            { "Vendors", GetVendors() }
        }),
        navTable = Table.ToNavigationTable(source, {"Name"}, "Name", "Data", "ItemKind", "ItemName", "IsLeaf")
    in
        navTable;


//AT.02252022 Navigation Table
Table.ToNavigationTable = (
    table as table,
    keyColumns as list,
    nameColumn as text,
    dataColumn as text,
    itemKindColumn as text,
    itemNameColumn as text,
    isLeafColumn as text
) as table =>
    let
        tableType = Value.Type(table),
        newTableType = Type.AddTableKey(tableType, keyColumns, true) meta 
        [
            NavigationTable.NameColumn = nameColumn, 
            NavigationTable.DataColumn = dataColumn,
            NavigationTable.ItemKindColumn = itemKindColumn, 
            Preview.DelayColumn = itemNameColumn, 
            NavigationTable.IsLeafColumn = isLeafColumn
        ],
        navigationTable = Value.ReplaceType(table, newTableType)
    in
        navigationTable;
/***************************************************Api Call Functions starts from here **************************************************/
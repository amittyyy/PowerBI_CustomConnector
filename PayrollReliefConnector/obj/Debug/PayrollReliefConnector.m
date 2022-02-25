// This file contains your Data Connector logic
section PayrollReliefConnector;

url = "https://dev-accountingexternalapi.accountingpower.com/api/customers";
//Payroll API OAuth2 required values
// client_id = "9KXdlkNoqxm%2B8PvzE%2FO87qzP3e294Cgy2fLo90Sd8sSllmMZQSQ6kVynpdXdcobE";
// client_secret = "d1nQHjTMs8iUmAte1AzQNaQedQ8k1rTlJVtoFMFHtVU=";
// 
//Ext Accounting power api client ID and Secrete
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

[DataSource.Kind="PayrollReliefConnector", Publish="PayrollReliefConnector.Publish"]
shared PayrollReliefConnector.Contents = () =>
    let 

        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                Headers = headers
            ]
        ))[customers],
      data = Table.FromList(dataList, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
      expandedColumns = Table.ExpandRecordColumn(data, "Column1", {"id", "name", "address1", "city", "state", "zip", "email", "phoneNumber", "balanceDue", "shipAddress1","shipCity", "shipState","shipZip","isInactive","salesTaxState", "salesTaxCounty", "salesTaxRate"})

    in
        expandedColumns;

// Data Source Kind description
PayrollReliefConnector = [
    TestConnection = (dataSourcePath) => { "PayrollReliefConnector.Contents", dataSourcePath },
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
PayrollReliefConnector.Publish = [
    Beta = true,
    Category = "Other",
    ButtonText = { Extension.LoadString("ButtonTitle"), Extension.LoadString("ButtonHelp") },
    LearnMoreUrl = "https://www.accountantsworld.com/",
    SourceImage = PayrollReliefConnector.Icons,
    SourceTypeImage = PayrollReliefConnector.Icons
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

PayrollReliefConnector.Icons = [
    Icon16 = { Extension.Contents("PayrollReliefConnector16.png"), Extension.Contents("PayrollReliefConnector20.png"), Extension.Contents("PayrollReliefConnector24.png"), Extension.Contents("PayrollReliefConnector32.png") },
    Icon32 = { Extension.Contents("PayrollReliefConnector32.png"), Extension.Contents("PayrollReliefConnector40.png"), Extension.Contents("PayrollReliefConnector48.png"), Extension.Contents("PayrollReliefConnector64.png") }
];

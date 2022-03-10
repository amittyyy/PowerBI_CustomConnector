/***************************************************************************************************************
**** AT. 020252022 Accoutnig power Customer Power BI Connector. ************************************************
**** This application will get data form Accoutning power External API via OAuth2 service. ********************* 

***** QA Links *************************************************************************************************
Links:
https://qa-accountingexternalapi.accountingpower.com/api/status
https://qa-accountingexternalapi.accountingpower.com/swagger/index.html

***************************************************************************************************************/

section AccountantsWorldConnector;

url = "https://qa-accountingexternalapi.accountingpower.com/api/";

//AT. 02252022 Ext Accounting power api client ID and Secrete
//Dev
// client_id = "36VXYCeMwpB0pHrmC6PZqkMdwu%2BYCCU%2FABmui1i60GjNOZ1oafQK9Zm5eKbOImoP";
// client_secret = "LY5q9u8CjNUQQp7t69N%2BLQvpGtxj69JTLEcIAvTOyEasnAF3GYWYS8aAaW6ElsRl";

//QA
client_id = "0eFqY55DvwJgZKLn3DNrRQt1ukq%2BDcEnxc%2FFxX%2Bcu1nWnbEH2RMrE1VC%2BGRtMABw";
client_secret = "sOhiOjjBVAgp7uZGUE8bN5XqigiYuadEVoCJ1oNg%2Fbc%2BaGo20ZcgeHsrPuayZoQS";

redirect_uri = "https://oauth.powerbi.com/views/oauthredirect.html";
token_uri = "https://qa-auth.accountantsoffice.com/connect/token";
authorize_uri = "https://qa-auth.accountantsoffice.com/connect/authorize";
logout_uri = "https://qa-auth.accountantsoffice.com/connect/logout";
// apiKey = "23911c167cc4417d940b4bb8f17c0a6d";

//AT. ACCounting client Api Keys
//Dev
// apiKey = "3911864bb1dd427a8983f2b10522f827";

//QA
apiKey = "dd0d3a174dbc4bf284b86b532eef3246";

//AT. Paramters for Transactions
fromCurrentDate= Number.ToText(Date.Year(DateTime.LocalNow())) & "-01-01";
toCurrentDate= Number.ToText(Date.Year(DateTime.LocalNow())) & "-12-31";
fromDateLastYear= Number.ToText(Date.Year(DateTime.LocalNow())-1) & "-01-01";
toDateLastYear= Number.ToText(Date.Year(DateTime.LocalNow())-1) & "-12-31";
emptyTable = Table.FromRecords({ [NoData=""] });

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
GetAccounts = () as table =>
        let 

        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                RelativePath = "accounts",
                Headers = headers
            ]
        )),
         accountsColumns = if  List.IsEmpty(dataList[accounts]) then
            emptyTable
       else  
            Table.ExpandRecordColumn(Table.FromList(dataList[accounts], Splitter.SplitByNothing(), null, null, ExtraValues.Error), "Column1", {"code", "description", "type", "typeDescription", "class", "classDescription", "category", "categoryDescription", "beginningBalance"})
          
    in
        accountsColumns;

GetCustomers = () as table =>
        let 

        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                RelativePath = "customers",
                Headers = headers
            ]
        )),
        customerColumns = if List.IsEmpty(dataList[customers]) then
            emptyTable
       else  
            Table.ExpandRecordColumn(Table.FromList(dataList[customers], Splitter.SplitByNothing(), null, null, ExtraValues.Error), "Column1", {"id", "name", "address1", "city", "state", "zip", "email", "phoneNumber", "balanceDue", "shipAddress1","shipCity", "shipState","shipZip","isInactive","salesTaxState", "salesTaxCounty", "salesTaxRate"})

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
        )),
         trialBalanceColumns = if List.IsEmpty(dataList[accounts]) then
            emptyTable
        else  
            Table.ExpandRecordColumn(Table.FromList(dataList[accounts], Splitter.SplitByNothing(), null, null, ExtraValues.Error), "Column1",{"code","type","description","beginningBalance","transactionsDebit","transactionsCredit","unadjustedBalance","adjustmentsDebit","adjustmentsCredit","adjustedBalance","notes"})

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
        )),
         vendorsColumns = if List.IsEmpty(dataList[vendors]) then
            emptyTable
        else  
            Table.ExpandRecordColumn(Table.FromList(dataList[vendors], Splitter.SplitByNothing(), null, null, ExtraValues.Error), "Column1",{"id","name","aliasName","is1099","address1","city","phoneNumber","faxNumber","isInactive","terms","defaultAccountCode","departmentId","isW9"})

    in
        vendorsColumns;

GetDepartments = () as table =>
        let 

        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                RelativePath = "departments",                
                Headers = headers
            ]
        )),
        departmentsColumns = if List.IsEmpty(dataList[departments]) then
            emptyTable
      else  
            Table.ExpandRecordColumn(Table.FromList(dataList[departments], Splitter.SplitByNothing(), null, null, ExtraValues.Error), "Column1",{"id","code","departmentName","shortName","isPrimary"})

    in
        departmentsColumns;

GetEmployees = () as table =>
        let 

        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                RelativePath = "employees",                
                Headers = headers
            ]
        )),
        employeesColumns = if List.IsEmpty(dataList[employees]) then
            emptyTable
      else
            Table.ExpandRecordColumn(Table.FromList(dataList[employees], Splitter.SplitByNothing(), null, null, ExtraValues.Error), "Column1",{"id","employeeNumber","fullName","firstName","lastName","companyPhone","state","departmentId","deptCode","deptName","salaryAccountCode","isFICAExempt","isCorporateOfficer","isSalesRep","isInactive"})
     
    in
        employeesColumns;

GetJobs = () as table =>
        let 

        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                RelativePath = "jobs",                
                Headers = headers
            ]
        )),
         jobssColumns = if List.IsEmpty(dataList[jobs]) then
            emptyTable
      else
            Table.ExpandRecordColumn(Table.FromList(dataList[jobs], Splitter.SplitByNothing(), null, null, ExtraValues.Error), "Column1",{"id","jobNumber","jobName","customerId","customerName","description","notes","startDate","projectedEndDate","status","jobTypeId","jobType","jobPhaseId","jobPhase","estimatedRevenue","estimatedExpenses","openingBalance"})
      
    in
        jobssColumns;

GetJobCategories = () as table =>
        let 

        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                RelativePath = "jobs/categories",                
                Headers = headers
            ]
        )),
        jobCategoriesColumns = if List.IsEmpty(dataList[list]) then
            emptyTable
      else    
            Table.ExpandRecordColumn(Table.FromList(dataList[list], Splitter.SplitByNothing(), null, null, ExtraValues.Error), "Column1",{"id","description"})     
    in
        jobCategoriesColumns;

GetPeriods = () as table =>
        let 

        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                RelativePath = "periods",                
                Headers = headers
            ]
        )),
         periodColumns = if List.IsEmpty(dataList[periods]) then
            emptyTable
        else    
            Table.FromList(dataList[periods], Splitter.SplitByNothing(), null, null, ExtraValues.Error)           
    in
        periodColumns;

GetPeriodAllowedForGeneralJournal = () as table =>
        let 

        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                RelativePath = "periods/AllowedForGeneralJournal",                
                Headers = headers
            ]
        )),
         jobPeriodAllowedColumns = if List.IsEmpty(dataList[periods]) then
            emptyTable
        else   
         Table.FromList(dataList[periods], Splitter.SplitByNothing(), null, null, ExtraValues.Error)           
    in
        jobPeriodAllowedColumns;

GetJobTypes = () as table =>
        let 

        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                RelativePath = "jobs/types",                
                Headers = headers
            ]
        )),
         jobTypesColumns = if List.IsEmpty(dataList[list]) then
            emptyTable
        else   
            Table.ExpandRecordColumn(Table.FromList(dataList[list], Splitter.SplitByNothing(), null, null, ExtraValues.Error), "Column1",{"id","description"}) 
     
    in
        jobTypesColumns;

GetJobPhases = () as table =>
        let 

        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                RelativePath = "jobs/phases",                
                Headers = headers
            ]
        )),
         jobPhasesColumns = if List.IsEmpty(dataList[list]) then
            emptyTable
        else   
            Table.ExpandRecordColumn(Table.FromList(dataList[list], Splitter.SplitByNothing(), null, null, ExtraValues.Error), "Column1",{"id","description"})     
    in
        jobPhasesColumns;

//AT. 02272022 Get Transaction of Journalcode = CD
GetTransactionsCDbyYeartoDate_LastYear = () as table =>
    
        let 
        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                RelativePath = "transactions/CD",  
                 Query = [fromDate= fromDateLastYear, toDate=toDateLastYear],
                Headers = headers
            ]
        )),
        
        transactionsCDColumns = if not (dataList[success]) then
            emptyTable
        else     
            Table.ExpandRecordColumn(Table.FromList(dataList[transactions], Splitter.SplitByNothing(), null, null, ExtraValues.Error), "Column1",{"id","groupId","journalCode","date","payerOrPayee","payerOrPayeeType","reference","accountCode","amount","debit","credit","accountCodeOffset","transactionType","cleared","is1099","memo","clearedDate","updatedBy"})
           
    in
        transactionsCDColumns;
GetTransactionsCDbyYeartoDate = () as table =>
    
        let 
        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                RelativePath = "transactions/CD",  
                 Query = [fromDate= fromCurrentDate, toDate=toCurrentDate],
                Headers = headers
            ]
        )),
        
        transactionsCDColumns = if not (dataList[success]) then
            emptyTable
        else     
            Table.ExpandRecordColumn(Table.FromList(dataList[transactions], Splitter.SplitByNothing(), null, null, ExtraValues.Error), "Column1",{"id","groupId","journalCode","date","payerOrPayee","payerOrPayeeType","reference","accountCode","amount","debit","credit","accountCodeOffset","transactionType","cleared","is1099","memo","clearedDate","updatedBy"})
           
    in
        transactionsCDColumns;

//AT. 02272022 Get Transaction of Journalcode = CR
GetTransactionsCRbyYeartoDate_LastYear = () as table =>
    
        let 
        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                RelativePath = "transactions/CR",  
                 Query = [fromDate= fromDateLastYear, toDate=toDateLastYear],
                Headers = headers
            ]
        )),
        
        transactionsCRColumns = if not (dataList[success]) then
            emptyTable
        else     
            Table.ExpandRecordColumn(Table.FromList(dataList[transactions], Splitter.SplitByNothing(), null, null, ExtraValues.Error), "Column1",{"id","groupId","journalCode","date","payerOrPayee","payerOrPayeeType","reference","accountCode","amount","debit","credit","accountCodeOffset","transactionType","cleared","is1099","memo","clearedDate","updatedBy"})
           
    in
        transactionsCRColumns;
GetTransactionsCRbyYeartoDate = () as table =>
    
        let 
        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                RelativePath = "transactions/CR",  
                 Query = [fromDate= fromCurrentDate, toDate=toCurrentDate],
                Headers = headers
            ]
        )),
        transactionsCRColumns = if not (dataList[success]) then
            emptyTable
        else     
            Table.ExpandRecordColumn(Table.FromList(dataList[transactions], Splitter.SplitByNothing(), null, null, ExtraValues.Error), "Column1",{"id","groupId","journalCode","date","payerOrPayee","payerOrPayeeType","reference","accountCode","amount","debit","credit","accountCodeOffset","transactionType","cleared","is1099","memo","clearedDate","updatedBy"})
           
    in
        transactionsCRColumns;

//AT. 02272022 Get Transaction of Journalcode = PR
GetTransactionsPRbyYeartoDate_LastYear = () as table =>
    
        let 
        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                RelativePath = "transactions/PR",  
                 Query = [fromDate= fromDateLastYear, toDate=toDateLastYear],
                Headers = headers
            ]
        )),

      transactionsPRColumns = if not (dataList[success]) then
            emptyTable
      else     
            Table.ExpandRecordColumn(Table.FromList(dataList[transactions], Splitter.SplitByNothing(), null, null, ExtraValues.Error), "Column1",{"id","groupId","journalCode","date","payerOrPayee","payerOrPayeeType","reference","accountCode","amount","debit","credit","accountCodeOffset","transactionType","cleared","is1099","memo","clearedDate","updatedBy"})

    in
        transactionsPRColumns;
GetTransactionsPRbyYeartoDate = () as table =>
    
        let 
        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                RelativePath = "transactions/PR",  
                 Query = [fromDate= fromCurrentDate, toDate=toCurrentDate],
                Headers = headers
            ]
        )),

      transactionsPRColumns = if not (dataList[success]) then
            emptyTable
      else     
            Table.ExpandRecordColumn(Table.FromList(dataList[transactions], Splitter.SplitByNothing(), null, null, ExtraValues.Error), "Column1",{"id","groupId","journalCode","date","payerOrPayee","payerOrPayeeType","reference","accountCode","amount","debit","credit","accountCodeOffset","transactionType","cleared","is1099","memo","clearedDate","updatedBy"})

    in
        transactionsPRColumns;

//AT. 02272022 Get Transaction of Journalcode = SJ
GetTransactionsSJbyYeartoDate_LastYear = () as table =>
    
        let 
        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                RelativePath = "transactions/SJ",  
                 Query = [fromDate= fromDateLastYear, toDate=toDateLastYear],
                Headers = headers
            ]
        )), 

      transactionsSJColumns = if not (dataList[success]) then
            emptyTable
      else     
            Table.ExpandRecordColumn(Table.FromList(dataList[transactions], Splitter.SplitByNothing(), null, null, ExtraValues.Error), "Column1",{"id","groupId","journalCode","date","payerOrPayee","payerOrPayeeType","reference","accountCode","amount","debit","credit","accountCodeOffset","transactionType","cleared","is1099","memo","clearedDate","updatedBy"})
     
    in
        transactionsSJColumns;
GetTransactionsSJbyYeartoDate = () as table =>
    
        let 
        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                RelativePath = "transactions/SJ",  
                 Query = [fromDate= fromCurrentDate, toDate=toCurrentDate],
                Headers = headers
            ]
        )), 

      transactionsSJColumns = if not (dataList[success]) then
            emptyTable
      else     
            Table.ExpandRecordColumn(Table.FromList(dataList[transactions], Splitter.SplitByNothing(), null, null, ExtraValues.Error), "Column1",{"id","groupId","journalCode","date","payerOrPayee","payerOrPayeeType","reference","accountCode","amount","debit","credit","accountCodeOffset","transactionType","cleared","is1099","memo","clearedDate","updatedBy"})
     
    in
        transactionsSJColumns;

 //AT. 02272022 Get Transaction of Journalcode = PJ
GetTransactionsPJbyYeartoDate_LastYear = () as table =>
    
        let 
        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                RelativePath = "transactions/PJ",  
                 Query = [fromDate= fromDateLastYear, toDate=toDateLastYear],
                Headers = headers
            ]
        )),

        transactionsPJColumns = if not (dataList[success]) then
            emptyTable
        else
            Table.ExpandRecordColumn(Table.FromList(dataList[transactions], Splitter.SplitByNothing(), null, null, ExtraValues.Error), "Column1",{"id","groupId","journalCode","date","payerOrPayee","payerOrPayeeType","reference","accountCode","amount","debit","credit","accountCodeOffset","transactionType","cleared","is1099","memo","clearedDate","updatedBy"})
     
    in
        transactionsPJColumns;
GetTransactionsPJbyYeartoDate = () as table =>
    
        let 
        headers = [#"x-api-key"=apiKey],
        dataList = Json.Document(Web.Contents(
            url,
            [
                RelativePath = "transactions/PJ",  
                 Query = [fromDate= fromCurrentDate, toDate=toCurrentDate],
                Headers = headers
            ]
        )),

        transactionsPJColumns = if not (dataList[success]) then
            emptyTable
        else
            Table.ExpandRecordColumn(Table.FromList(dataList[transactions], Splitter.SplitByNothing(), null, null, ExtraValues.Error), "Column1",{"id","groupId","journalCode","date","payerOrPayee","payerOrPayeeType","reference","accountCode","amount","debit","credit","accountCodeOffset","transactionType","cleared","is1099","memo","clearedDate","updatedBy"})
     
    in
        transactionsPJColumns;

//AT. 02252022 Combining all api return table into one table.
AccountingTables = () as table =>
    let      

        source = #table({"Name", "Data"}, {
            { "Customers", GetCustomers() },
            { "TrailBalances", GetTrailBalance() },
            { "Vendors", GetVendors() },
            {"Accounts", GetAccounts() },
            {"Departments", GetDepartments()},
                {"Employees", GetEmployees()},
                {"Jobs", GetJobs()},
                {"JobCatogories", GetJobCategories()},
                {"JobTypes", GetJobTypes()},
                {"JobPhases", GetJobPhases()},
            {"Periods", GetPeriods()},
            {"PeriodsAllowedForGeneralJournal", GetPeriodAllowedForGeneralJournal()},
            {"CashDisbursementTransactions_LastYear_YearToDate", GetTransactionsCDbyYeartoDate_LastYear()},
            {"CashDisbursementTransactions_CurrentYear_YearToDate", GetTransactionsCDbyYeartoDate()},
            {"CashReceiptsTransactions_LastYear_YearToDate", GetTransactionsCRbyYeartoDate_LastYear()},
                {"CashReceiptsTransactions_CurrentYear_YearToDate", GetTransactionsCRbyYeartoDate()},
                {"PayrollTransactions_LastYear_YearToDate", GetTransactionsPRbyYeartoDate_LastYear()},
                {"PayrollTransactions_CurrentYear_YearToDate", GetTransactionsPRbyYeartoDate()},
                {"SalesTransactions_LastYear_YearToDate", GetTransactionsSJbyYeartoDate_LastYear()},
                {"SalesTransactions_CurrentYear_YearToDate", GetTransactionsSJbyYeartoDate()},
            {"PurchaseTransactions_LastYear_YearToDate", GetTransactionsPJbyYeartoDate_LastYear()},
            {"PurchaseTransactions_CurrentYear_YearToDate", GetTransactionsPJbyYeartoDate()}
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
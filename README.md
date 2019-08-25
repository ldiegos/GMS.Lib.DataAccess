# Introduction
Library to help using connections and querys to MSSQLServer, SQLite and MSAccess.
With this lib, you will focus your efforts in create the application and ignore the connection and methods to access the database.

# Getting Started
* [SQLite management](#sqlite-management): Know the methds and how to use it.
* [SQLServer management](#sqlserver-management): Know the methds and how to use it.
* [MSAccess management](#msaccess-management): Know the methds and how to use it.


# Contribute
TODO: Explain how other users and developers can contribute to make your code better. 

If you want to learn more about creating good readme files then refer the following [guidelines](https://www.visualstudio.com/en-us/docs/git/create-a-readme). You can also seek inspiration from the below readme files:
- [ASP.NET Core](https://github.com/aspnet/Home)
- [Visual Studio Code](https://github.com/Microsoft/vscode)
- [Chakra Core](https://github.com/Microsoft/ChakraCore)

## SQLite management
### Introduction
SQLite is a portable database, very usefull when you have limited resources and installation permissions.
### Instance
The best way to use the capabilities is with the interface layer: **_IGmsLibDataAccess_ objDataBaseAccess** and then using the class constructor: **_new GmsLibDataAcessSqlite(StrConnectionString, out result);_**

Using the interface will be very helpfull if you plan to migrate the SQLite database to and SQLServer database. You will only need to change the class constructor name.

### Public methods
**Constructors:**
* **public GmsLibDataAcessSqlite(string strConnectionStringBd, out bool result)**: With the connection string the method will check if the database exists and create it. 
It will use the default connection parameters to configure the object:
   - Max connection retries = 3
   - Wait seconds between reconnections = 10
   - Max execution retries = 10
   - Wait seconds between executions = 6
* **public GmsLibDataAcessSqlite(string connectionString, int numMaxReconnections, int numSecondsBetweenReconnection
            , int maxExecutionRetries, int numSecondsBetweenExecutionRetries, out bool bolFnResult)**: With the connection string the method will check if the database exists and create it. 
It will use the received connection parameters to configure the object.

Examples:

string strSQLConnection = "data source=#DIRECTORY#\#DATABASE#.#EXTENSION#;Compress=True;synchronous=Off;foreign keys=True;pooling=True;Journal Mode=Wal;Default Isolation Level=ReadCommitted";
* #DIRECTORY# is the fullpath to the database that you want
* #DATABASE# is the name of the database.
* #EXTENSION# is a extension to the sqlite database. For example db, sqlite, sqlite2 or sqlite3, but any other extension will work.

var sqliteconnect = new GmsLibDataAcessSqlite(strSQLConnection, out result)
Console.WriteLine(sqliteconnect.GetError);

var sqliteconnect = new GmsLibDataAcessSqlite(strSQLConnection, 10, 5, 3, 1, out result)
Console.WriteLine(sqliteconnect.GetError);

**Connections:**
* **public bool OpenSqlConnection()**: This method will open the connection and leave it open until the object is dispose or close it with the CloseSqlConnection()
* **public bool CloseSqlConnection()**: This method will close the previously opened connection. 

Both methods will return a true or false with the result of the method. If false is return, please, check the _object_.GetError to know the error message.



## SQLServer management


## MSAccess management

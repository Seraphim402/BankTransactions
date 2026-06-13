using Microsoft.Data.Sqlite;

namespace BankTransactions.Core;

public sealed class BankDatabase
{
    private readonly string _connectionString;

    public BankDatabase(string databasePath)
    {
        var builder = new SqliteConnectionStringBuilder { DataSource = databasePath };
        _connectionString = builder.ToString();
    }

    public SqliteConnection OpenConnection()
    {
        var connection = new SqliteConnection(_connectionString);
        connection.Open();
        using var command = connection.CreateCommand();
        command.CommandText = "PRAGMA foreign_keys = ON";
        command.ExecuteNonQuery();
        return connection;
    }

    public void Initialize()
    {
        using var connection = OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText =
            """
            PRAGMA foreign_keys = ON;

            CREATE TABLE IF NOT EXISTS Users (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                Login TEXT NOT NULL UNIQUE,
                Password TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS Clients (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                Type INTEGER NOT NULL,
                Name TEXT NOT NULL,
                TaxNumber TEXT NOT NULL,
                Phone TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS Accounts (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                ClientId INTEGER NOT NULL,
                Type INTEGER NOT NULL,
                Currency TEXT NOT NULL,
                Balance TEXT NOT NULL,
                CreatedAt TEXT NOT NULL,
                FOREIGN KEY (ClientId) REFERENCES Clients(Id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS PartnerBanks (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                Name TEXT NOT NULL,
                Country TEXT NOT NULL,
                IsForeign INTEGER NOT NULL
            );

            CREATE TABLE IF NOT EXISTS Commissions (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                AccountType INTEGER NOT NULL,
                IsForeignPartner INTEGER NOT NULL,
                Percent TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS Transactions (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                Type INTEGER NOT NULL,
                FromAccountId INTEGER NOT NULL,
                ToAccountId INTEGER,
                PartnerBankId INTEGER,
                Amount TEXT NOT NULL,
                Commission TEXT NOT NULL,
                CreatedAt TEXT NOT NULL,
                Comment TEXT NOT NULL,
                FOREIGN KEY (FromAccountId) REFERENCES Accounts(Id) ON DELETE CASCADE,
                FOREIGN KEY (ToAccountId) REFERENCES Accounts(Id) ON DELETE CASCADE,
                FOREIGN KEY (PartnerBankId) REFERENCES PartnerBanks(Id)
            );

            INSERT OR IGNORE INTO Users (Login, Password) VALUES ('admin', 'admin');

            INSERT OR IGNORE INTO PartnerBanks (Id, Name, Country, IsForeign) VALUES
                (1, 'Городской партнер', 'Россия', 0),
                (2, 'Euro Partner Bank', 'Германия', 1),
                (3, 'Asia Trade Bank', 'Китай', 1);

            INSERT OR IGNORE INTO Commissions (Id, AccountType, IsForeignPartner, Percent) VALUES
                (1, 1, 0, '0.5'),
                (2, 1, 1, '1.5'),
                (3, 2, 0, '1.0'),
                (4, 2, 1, '2.0'),
                (5, 3, 0, '0.2'),
                (6, 3, 1, '1.0');
            """;
        command.ExecuteNonQuery();
    }
}

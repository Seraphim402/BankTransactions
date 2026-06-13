using Microsoft.Data.Sqlite;
using System.Globalization;

namespace BankTransactions.Core;

public sealed class BankService
{
    private readonly BankDatabase _database;

    public BankService(BankDatabase database)
    {
        _database = database;
    }

    public bool Login(string login, string password)
    {
        using var connection = _database.OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT COUNT(*) FROM Users WHERE Login = $login AND Password = $password";
        command.Parameters.AddWithValue("$login", login);
        command.Parameters.AddWithValue("$password", password);
        return Convert.ToInt32(command.ExecuteScalar()) == 1;
    }

    public int CreateClient(CreateClientRequest request)
    {
        using var connection = _database.OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText =
            """
            INSERT INTO Clients (Type, Name, TaxNumber, Phone)
            VALUES ($type, $name, $taxNumber, $phone);
            SELECT last_insert_rowid();
            """;
        command.Parameters.AddWithValue("$type", (int)request.Type);
        command.Parameters.AddWithValue("$name", request.Name);
        command.Parameters.AddWithValue("$taxNumber", request.TaxNumber);
        command.Parameters.AddWithValue("$phone", request.Phone);
        return Convert.ToInt32(command.ExecuteScalar());
    }

    public void DeleteClient(int id)
    {
        using var connection = _database.OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText = "DELETE FROM Clients WHERE Id = $id";
        command.Parameters.AddWithValue("$id", id);
        command.ExecuteNonQuery();
    }

    public IReadOnlyList<Client> GetClients()
    {
        using var connection = _database.OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT Id, Type, Name, TaxNumber, Phone FROM Clients ORDER BY Id";
        using var reader = command.ExecuteReader();
        var clients = new List<Client>();
        while (reader.Read())
        {
            clients.Add(new Client(
                reader.GetInt32(0),
                (ClientType)reader.GetInt32(1),
                reader.GetString(2),
                reader.GetString(3),
                reader.GetString(4)));
        }

        return clients;
    }

    public int CreateAccount(CreateAccountRequest request)
    {
        using var connection = _database.OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText =
            """
            INSERT INTO Accounts (ClientId, Type, Currency, Balance, CreatedAt)
            VALUES ($clientId, $type, $currency, '0', $createdAt);
            SELECT last_insert_rowid();
            """;
        command.Parameters.AddWithValue("$clientId", request.ClientId);
        command.Parameters.AddWithValue("$type", (int)request.Type);
        command.Parameters.AddWithValue("$currency", request.Currency.ToUpperInvariant());
        command.Parameters.AddWithValue("$createdAt", DateTime.UtcNow.ToString("O"));
        return Convert.ToInt32(command.ExecuteScalar());
    }

    public IReadOnlyList<Account> GetAccounts()
    {
        using var connection = _database.OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT Id, ClientId, Type, Currency, Balance, CreatedAt FROM Accounts ORDER BY Id";
        using var reader = command.ExecuteReader();
        var accounts = new List<Account>();
        while (reader.Read())
        {
            accounts.Add(new Account(
                reader.GetInt32(0),
                reader.GetInt32(1),
                (AccountType)reader.GetInt32(2),
                reader.GetString(3),
                ParseDecimal(reader.GetString(4)),
                DateTime.Parse(reader.GetString(5))));
        }

        return accounts;
    }

    public IReadOnlyList<PartnerBank> GetPartnerBanks()
    {
        using var connection = _database.OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT Id, Name, Country, IsForeign FROM PartnerBanks ORDER BY Id";
        using var reader = command.ExecuteReader();
        var banks = new List<PartnerBank>();
        while (reader.Read())
        {
            banks.Add(new PartnerBank(reader.GetInt32(0), reader.GetString(1), reader.GetString(2), reader.GetInt32(3) == 1));
        }

        return banks;
    }

    public int Deposit(DepositRequest request)
    {
        if (request.Amount <= 0)
        {
            throw new InvalidOperationException("Сумма пополнения должна быть больше нуля.");
        }

        using var connection = _database.OpenConnection();
        using var transaction = connection.BeginTransaction();
        var balance = GetBalance(connection, request.AccountId);
        SetBalance(connection, request.AccountId, balance + request.Amount);
        var id = InsertTransaction(connection, TransactionType.Deposit, request.AccountId, null, null, request.Amount, 0, request.Comment);
        transaction.Commit();
        return id;
    }

    public int Transfer(TransferRequest request)
    {
        if (request.Amount <= 0)
        {
            throw new InvalidOperationException("Сумма перевода должна быть больше нуля.");
        }

        using var connection = _database.OpenConnection();
        using var transaction = connection.BeginTransaction();

        var from = GetAccount(connection, request.FromAccountId);
        _ = GetAccount(connection, request.ToAccountId);
        var partner = GetPartnerBank(connection, request.PartnerBankId);
        var commission = CalculateCommission(connection, from.Type, partner.IsForeign, request.Amount);
        var total = request.Amount + commission;

        if (from.Balance < total)
        {
            throw new InvalidOperationException("Недостаточно средств с учетом комиссии.");
        }

        var toBalance = GetBalance(connection, request.ToAccountId);
        SetBalance(connection, request.FromAccountId, from.Balance - total);
        SetBalance(connection, request.ToAccountId, toBalance + request.Amount);
        var id = InsertTransaction(connection, TransactionType.Transfer, request.FromAccountId, request.ToAccountId, request.PartnerBankId, request.Amount, commission, request.Comment);

        transaction.Commit();
        return id;
    }

    public IReadOnlyList<BankTransaction> GetTransactions()
    {
        using var connection = _database.OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText =
            """
            SELECT Id, Type, FromAccountId, ToAccountId, PartnerBankId, Amount, Commission, CreatedAt, Comment
            FROM Transactions
            ORDER BY Id DESC
            """;
        using var reader = command.ExecuteReader();
        var transactions = new List<BankTransaction>();
        while (reader.Read())
        {
            transactions.Add(new BankTransaction(
                reader.GetInt32(0),
                (TransactionType)reader.GetInt32(1),
                reader.GetInt32(2),
                reader.IsDBNull(3) ? null : reader.GetInt32(3),
                reader.IsDBNull(4) ? null : reader.GetInt32(4),
                ParseDecimal(reader.GetString(5)),
                ParseDecimal(reader.GetString(6)),
                DateTime.Parse(reader.GetString(7)),
                reader.GetString(8)));
        }

        return transactions;
    }

    private static Account GetAccount(SqliteConnection connection, int accountId)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT Id, ClientId, Type, Currency, Balance, CreatedAt FROM Accounts WHERE Id = $id";
        command.Parameters.AddWithValue("$id", accountId);
        using var reader = command.ExecuteReader();
        if (!reader.Read())
        {
            throw new InvalidOperationException("Счет не найден.");
        }

        return new Account(
            reader.GetInt32(0),
            reader.GetInt32(1),
            (AccountType)reader.GetInt32(2),
            reader.GetString(3),
            ParseDecimal(reader.GetString(4)),
            DateTime.Parse(reader.GetString(5)));
    }

    private static PartnerBank GetPartnerBank(SqliteConnection connection, int bankId)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT Id, Name, Country, IsForeign FROM PartnerBanks WHERE Id = $id";
        command.Parameters.AddWithValue("$id", bankId);
        using var reader = command.ExecuteReader();
        if (!reader.Read())
        {
            throw new InvalidOperationException("Банк-партнер не найден.");
        }

        return new PartnerBank(reader.GetInt32(0), reader.GetString(1), reader.GetString(2), reader.GetInt32(3) == 1);
    }

    private static decimal GetBalance(SqliteConnection connection, int accountId) => GetAccount(connection, accountId).Balance;

    private static void SetBalance(SqliteConnection connection, int accountId, decimal balance)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "UPDATE Accounts SET Balance = $balance WHERE Id = $id";
        command.Parameters.AddWithValue("$balance", FormatDecimal(balance));
        command.Parameters.AddWithValue("$id", accountId);
        command.ExecuteNonQuery();
    }

    private static decimal CalculateCommission(SqliteConnection connection, AccountType accountType, bool isForeign, decimal amount)
    {
        using var command = connection.CreateCommand();
        command.CommandText =
            """
            SELECT Percent FROM Commissions
            WHERE AccountType = $accountType AND IsForeignPartner = $isForeign
            """;
        command.Parameters.AddWithValue("$accountType", (int)accountType);
        command.Parameters.AddWithValue("$isForeign", isForeign ? 1 : 0);
        var percent = ParseDecimal((string)command.ExecuteScalar()!);
        return Math.Round(amount * percent / 100m, 2);
    }

    private static int InsertTransaction(
        SqliteConnection connection,
        TransactionType type,
        int fromAccountId,
        int? toAccountId,
        int? partnerBankId,
        decimal amount,
        decimal commission,
        string comment)
    {
        using var command = connection.CreateCommand();
        command.CommandText =
            """
            INSERT INTO Transactions (Type, FromAccountId, ToAccountId, PartnerBankId, Amount, Commission, CreatedAt, Comment)
            VALUES ($type, $fromAccountId, $toAccountId, $partnerBankId, $amount, $commission, $createdAt, $comment);
            SELECT last_insert_rowid();
            """;
        command.Parameters.AddWithValue("$type", (int)type);
        command.Parameters.AddWithValue("$fromAccountId", fromAccountId);
        command.Parameters.AddWithValue("$toAccountId", (object?)toAccountId ?? DBNull.Value);
        command.Parameters.AddWithValue("$partnerBankId", (object?)partnerBankId ?? DBNull.Value);
        command.Parameters.AddWithValue("$amount", FormatDecimal(amount));
        command.Parameters.AddWithValue("$commission", FormatDecimal(commission));
        command.Parameters.AddWithValue("$createdAt", DateTime.UtcNow.ToString("O"));
        command.Parameters.AddWithValue("$comment", comment);
        return Convert.ToInt32(command.ExecuteScalar());
    }

    private static decimal ParseDecimal(string value) =>
        decimal.Parse(value, CultureInfo.InvariantCulture);

    private static string FormatDecimal(decimal value) =>
        value.ToString(CultureInfo.InvariantCulture);
}

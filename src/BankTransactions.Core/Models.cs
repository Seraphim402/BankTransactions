namespace BankTransactions.Core;

public enum ClientType
{
    Individual = 1,
    LegalEntity = 2
}

public enum AccountType
{
    Salary = 1,
    Currency = 2,
    Savings = 3
}

public enum TransactionType
{
    Deposit = 1,
    Transfer = 2
}

public sealed record Client(int Id, ClientType Type, string Name, string TaxNumber, string Phone);

public sealed record Account(
    int Id,
    int ClientId,
    AccountType Type,
    string Currency,
    decimal Balance,
    DateTime CreatedAt);

public sealed record PartnerBank(int Id, string Name, string Country, bool IsForeign);

public sealed record CommissionRule(int Id, AccountType AccountType, bool IsForeignPartner, decimal Percent);

public sealed record BankTransaction(
    int Id,
    TransactionType Type,
    int FromAccountId,
    int? ToAccountId,
    int? PartnerBankId,
    decimal Amount,
    decimal Commission,
    DateTime CreatedAt,
    string Comment);

public sealed record LoginRequest(string Login, string Password);

public sealed record CreateClientRequest(ClientType Type, string Name, string TaxNumber, string Phone);

public sealed record CreateAccountRequest(int ClientId, AccountType Type, string Currency);

public sealed record DepositRequest(int AccountId, decimal Amount, string Comment);

public sealed record TransferRequest(
    int FromAccountId,
    int ToAccountId,
    int PartnerBankId,
    decimal Amount,
    string Comment);

public sealed record BankRequest(string Command, string? Token, string? Payload);

public sealed record BankResponse(bool Success, string Message, string? Payload = null);

using BankTransactions.Core;
using Xunit;

namespace BankTransactions.Tests;

public sealed class BankServiceTests
{
    [Fact]
    public void Login_ReturnsTrue_ForDefaultAdmin()
    {
        var service = CreateService();

        var result = service.Login("admin", "admin");

        Assert.True(result);
    }

    [Fact]
    public void Deposit_IncreasesAccountBalance()
    {
        var service = CreateService();
        var clientId = service.CreateClient(new CreateClientRequest(ClientType.Individual, "Иван Иванов", "123", "+79990000000"));
        var accountId = service.CreateAccount(new CreateAccountRequest(clientId, AccountType.Salary, "RUB"));

        service.Deposit(new DepositRequest(accountId, 1000m, "Тестовое пополнение"));

        var account = Assert.Single(service.GetAccounts());
        Assert.Equal(1000m, account.Balance);
    }

    [Fact]
    public void Transfer_WithdrawsAmountAndCommission()
    {
        var service = CreateService();
        var firstClientId = service.CreateClient(new CreateClientRequest(ClientType.Individual, "Иван Иванов", "123", "+79990000000"));
        var secondClientId = service.CreateClient(new CreateClientRequest(ClientType.LegalEntity, "ООО Ромашка", "456", "+79991111111"));
        var fromAccountId = service.CreateAccount(new CreateAccountRequest(firstClientId, AccountType.Salary, "RUB"));
        var toAccountId = service.CreateAccount(new CreateAccountRequest(secondClientId, AccountType.Currency, "EUR"));

        service.Deposit(new DepositRequest(fromAccountId, 1000m, "Начальный баланс"));
        service.Transfer(new TransferRequest(fromAccountId, toAccountId, 2, 100m, "Международный перевод"));

        var accounts = service.GetAccounts();
        var from = accounts.Single(account => account.Id == fromAccountId);
        var to = accounts.Single(account => account.Id == toAccountId);
        var transfer = service.GetTransactions().First(transaction => transaction.Type == TransactionType.Transfer);

        Assert.Equal(898.50m, from.Balance);
        Assert.Equal(100m, to.Balance);
        Assert.Equal(1.50m, transfer.Commission);
    }

    [Fact]
    public void Transfer_Throws_WhenBalanceIsNotEnough()
    {
        var service = CreateService();
        var firstClientId = service.CreateClient(new CreateClientRequest(ClientType.Individual, "Иван Иванов", "123", "+79990000000"));
        var secondClientId = service.CreateClient(new CreateClientRequest(ClientType.Individual, "Петр Петров", "456", "+79991111111"));
        var fromAccountId = service.CreateAccount(new CreateAccountRequest(firstClientId, AccountType.Salary, "RUB"));
        var toAccountId = service.CreateAccount(new CreateAccountRequest(secondClientId, AccountType.Salary, "RUB"));

        var ex = Assert.Throws<InvalidOperationException>(() =>
            service.Transfer(new TransferRequest(fromAccountId, toAccountId, 1, 100m, "Без денег")));

        Assert.Contains("Недостаточно средств", ex.Message);
    }

    private static BankService CreateService()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bank-tests-{Guid.NewGuid():N}.db");
        var database = new BankDatabase(path);
        database.Initialize();
        return new BankService(database);
    }
}

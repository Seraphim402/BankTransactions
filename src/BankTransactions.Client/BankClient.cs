using System.IO;
using System.Net.Sockets;
using System.Text.Json;
using BankTransactions.Core;

namespace BankTransactions.Client;

public sealed class BankClient : IAsyncDisposable
{
    private TcpClient? _client;
    private StreamReader? _reader;
    private StreamWriter? _writer;
    private string? _token;

    public async Task ConnectAsync()
    {
        if (_client?.Connected == true)
        {
            return;
        }

        _client = new TcpClient();
        await _client.ConnectAsync("127.0.0.1", 5050);
        var stream = _client.GetStream();
        _reader = new StreamReader(stream);
        _writer = new StreamWriter(stream) { AutoFlush = true };
    }

    public async Task LoginAsync(string login, string password)
    {
        var token = await SendAsync<string>("Login", new LoginRequest(login, password));
        _token = token;
    }

    public Task<IReadOnlyList<BankTransactions.Core.Client>> GetClientsAsync() =>
        SendAsync<IReadOnlyList<BankTransactions.Core.Client>>("ListClients", null);

    public Task<int> CreateClientAsync(CreateClientRequest request) =>
        SendAsync<int>("CreateClient", request);

    public Task DeleteClientAsync(int id) =>
        SendAsync<object>("DeleteClient", id);

    public Task<IReadOnlyList<Account>> GetAccountsAsync() =>
        SendAsync<IReadOnlyList<Account>>("ListAccounts", null);

    public Task<int> CreateAccountAsync(CreateAccountRequest request) =>
        SendAsync<int>("CreateAccount", request);

    public Task<int> DepositAsync(DepositRequest request) =>
        SendAsync<int>("Deposit", request);

    public Task<int> TransferAsync(TransferRequest request) =>
        SendAsync<int>("Transfer", request);

    public Task<IReadOnlyList<BankTransaction>> GetTransactionsAsync() =>
        SendAsync<IReadOnlyList<BankTransaction>>("ListTransactions", null);

    public Task<IReadOnlyList<PartnerBank>> GetPartnerBanksAsync() =>
        SendAsync<IReadOnlyList<PartnerBank>>("ListPartnerBanks", null);

    private async Task<T> SendAsync<T>(string command, object? payload)
    {
        await ConnectAsync();
        var request = new BankRequest(command, _token, payload is null ? null : JsonSerializer.Serialize(payload));
        await _writer!.WriteLineAsync(JsonSerializer.Serialize(request));
        var line = await _reader!.ReadLineAsync() ?? throw new InvalidOperationException("Сервер закрыл соединение.");
        var response = JsonSerializer.Deserialize<BankResponse>(line) ?? throw new InvalidOperationException("Пустой ответ сервера.");
        if (!response.Success)
        {
            throw new InvalidOperationException(response.Message);
        }

        if (typeof(T) == typeof(object))
        {
            return default!;
        }

        return JsonSerializer.Deserialize<T>(response.Payload ?? "null")
            ?? throw new InvalidOperationException("Не удалось прочитать ответ сервера.");
    }

    public async ValueTask DisposeAsync()
    {
        if (_writer is not null)
        {
            await _writer.DisposeAsync();
        }

        _reader?.Dispose();
        _client?.Dispose();
    }
}

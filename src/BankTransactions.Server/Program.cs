using System.Collections.Concurrent;
using System.Net;
using System.Net.Sockets;
using System.Text.Json;
using BankTransactions.Core;

const int port = 5050;

var databasePath = Path.Combine(AppContext.BaseDirectory, "bank.db");
var database = new BankDatabase(databasePath);
database.Initialize();
var service = new BankService(database);
var sessions = new ConcurrentDictionary<string, TcpClient>();
var listener = new TcpListener(IPAddress.Loopback, port);

listener.Start();
Console.WriteLine($"Сервер банковских транзакций запущен на порту {port}.");
Console.WriteLine($"Файл базы данных: {databasePath}");
Console.WriteLine("Ожидание подключений клиентов...");

while (true)
{
    var client = await listener.AcceptTcpClientAsync();
    var sessionId = Guid.NewGuid().ToString("N");
    sessions[sessionId] = client;
    PrintClients();
    _ = Task.Run(() => HandleClientAsync(sessionId, client));
}

async Task HandleClientAsync(string sessionId, TcpClient client)
{
    await using var stream = client.GetStream();
    using var reader = new StreamReader(stream);
    await using var writer = new StreamWriter(stream) { AutoFlush = true };
    var token = string.Empty;

    try
    {
        while (await reader.ReadLineAsync() is { } line)
        {
            var request = JsonSerializer.Deserialize<BankRequest>(line);
            if (request is null)
            {
                await SendAsync(writer, new BankResponse(false, "Пустой запрос."));
                continue;
            }

            var response = Execute(request, token, out var newToken);
            token = newToken;
            await SendAsync(writer, response);
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Клиент {sessionId} отключился с ошибкой: {ex.Message}");
    }
    finally
    {
        sessions.TryRemove(sessionId, out _);
        client.Dispose();
        PrintClients();
    }
}

BankResponse Execute(BankRequest request, string currentToken, out string newToken)
{
    newToken = currentToken;

    try
    {
        switch (request.Command)
        {
            case "Login":
                var login = ReadPayload<LoginRequest>(request);
                if (!service.Login(login.Login, login.Password))
                {
                    return new BankResponse(false, "Неверный логин или пароль.");
                }

                newToken = Guid.NewGuid().ToString("N");
                return new BankResponse(true, "Вход выполнен.", JsonSerializer.Serialize(newToken));

            case "ListClients":
                return Json(service.GetClients());

            case "CreateClient":
                RequireLogin(request, currentToken);
                return Json(service.CreateClient(ReadPayload<CreateClientRequest>(request)), "Клиент создан.");

            case "DeleteClient":
                RequireLogin(request, currentToken);
                service.DeleteClient(ReadPayload<int>(request));
                return new BankResponse(true, "Клиент удален.");

            case "ListAccounts":
                return Json(service.GetAccounts());

            case "CreateAccount":
                RequireLogin(request, currentToken);
                return Json(service.CreateAccount(ReadPayload<CreateAccountRequest>(request)), "Счет создан.");

            case "Deposit":
                RequireLogin(request, currentToken);
                return Json(service.Deposit(ReadPayload<DepositRequest>(request)), "Счет пополнен.");

            case "Transfer":
                RequireLogin(request, currentToken);
                return Json(service.Transfer(ReadPayload<TransferRequest>(request)), "Перевод выполнен.");

            case "ListTransactions":
                return Json(service.GetTransactions());

            case "ListPartnerBanks":
                return Json(service.GetPartnerBanks());

            default:
                return new BankResponse(false, "Неизвестная команда.");
        }
    }
    catch (Exception ex)
    {
        return new BankResponse(false, ex.Message);
    }
}

static void RequireLogin(BankRequest current, string currentToken)
{
    if (string.IsNullOrWhiteSpace(current.Token) || current.Token != currentToken)
    {
        throw new InvalidOperationException("Сначала выполните вход.");
    }
}

T ReadPayload<T>(BankRequest request)
{
    if (string.IsNullOrWhiteSpace(request.Payload))
    {
        throw new InvalidOperationException("Нет данных для команды.");
    }

    return JsonSerializer.Deserialize<T>(request.Payload)
        ?? throw new InvalidOperationException("Не удалось прочитать данные команды.");
}

BankResponse Json<T>(T payload, string message = "Операция выполнена.") =>
    new(true, message, JsonSerializer.Serialize(payload));

static Task SendAsync(StreamWriter writer, BankResponse response) =>
    writer.WriteLineAsync(JsonSerializer.Serialize(response));

void PrintClients()
{
    Console.WriteLine($"Подключенных клиентов: {sessions.Count}");
}

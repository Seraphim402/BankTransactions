using System.Windows;
using System.Windows.Controls;
using BankTransactions.Core;

namespace BankTransactions.Client;

public partial class MainWindow : Window
{
    private readonly BankClient _client = new();

    public MainWindow()
    {
        InitializeComponent();
        Loaded += async (_, _) => await RefreshAllAsync();
        Closed += async (_, _) => await _client.DisposeAsync();
    }

    private async void LoginClick(object sender, RoutedEventArgs e) =>
        await RunAsync(async () =>
        {
            await _client.LoginAsync(LoginBox.Text, PasswordBox.Password);
            StatusText.Text = "Вход выполнен.";
            await RefreshAllAsync();
        });

    private async void CreateClientClick(object sender, RoutedEventArgs e) =>
        await RunAsync(async () =>
        {
            await _client.CreateClientAsync(new CreateClientRequest(
                GetEnum<ClientType>(ClientTypeBox),
                ClientNameBox.Text,
                TaxNumberBox.Text,
                PhoneBox.Text));
            await RefreshAllAsync();
        });

    private async void DeleteClientClick(object sender, RoutedEventArgs e) =>
        await RunAsync(async () =>
        {
            if (ClientsGrid.SelectedItem is not BankTransactions.Core.Client client)
            {
                throw new InvalidOperationException("Выберите клиента.");
            }

            await _client.DeleteClientAsync(client.Id);
            await RefreshAllAsync();
        });

    private async void CreateAccountClick(object sender, RoutedEventArgs e) =>
        await RunAsync(async () =>
        {
            await _client.CreateAccountAsync(new CreateAccountRequest(
                int.Parse(AccountClientIdBox.Text),
                GetEnum<AccountType>(AccountTypeBox),
                CurrencyBox.Text));
            await RefreshAllAsync();
        });

    private async void DepositClick(object sender, RoutedEventArgs e) =>
        await RunAsync(async () =>
        {
            await _client.DepositAsync(new DepositRequest(
                int.Parse(DepositAccountIdBox.Text),
                decimal.Parse(DepositAmountBox.Text),
                "Пополнение через кассу"));
            await RefreshAllAsync();
        });

    private async void TransferClick(object sender, RoutedEventArgs e) =>
        await RunAsync(async () =>
        {
            await _client.TransferAsync(new TransferRequest(
                int.Parse(FromAccountBox.Text),
                int.Parse(ToAccountBox.Text),
                (int)PartnerBankBox.SelectedValue,
                decimal.Parse(TransferAmountBox.Text),
                TransferCommentBox.Text));
            await RefreshAllAsync();
        });

    private async Task RefreshAllAsync()
    {
        try
        {
            ClientsGrid.ItemsSource = await _client.GetClientsAsync();
            AccountsGrid.ItemsSource = await _client.GetAccountsAsync();
            TransactionsGrid.ItemsSource = await _client.GetTransactionsAsync();
            var banks = await _client.GetPartnerBanksAsync();
            PartnerBankBox.ItemsSource = banks;
            if (PartnerBankBox.SelectedIndex < 0 && banks.Count > 0)
            {
                PartnerBankBox.SelectedIndex = 0;
            }
        }
        catch
        {
            StatusText.Text = "Запустите сервер перед работой с окном.";
        }
    }

    private async Task RunAsync(Func<Task> action)
    {
        try
        {
            await action();
            StatusText.Text = "Готово.";
        }
        catch (Exception ex)
        {
            StatusText.Text = ex.Message;
        }
    }

    private static T GetEnum<T>(ComboBox comboBox) where T : struct
    {
        var selected = (ComboBoxItem)comboBox.SelectedItem;
        return Enum.Parse<T>((string)selected.Tag);
    }
}

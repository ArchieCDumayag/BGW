using BillingSystem.Models;

namespace BillingSystem.Services;

public interface IAuthService
{
    Task<UserAccount?> ValidateAsync(string username, string password, string role);
}

public sealed class AuthService(IBillingStore store) : IAuthService
{
    public async Task<UserAccount?> ValidateAsync(string username, string password, string role)
    {
        if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
        {
            return null;
        }

        var data = await store.GetAsync();
        var account = data.UserAccounts.FirstOrDefault(user =>
            user.IsActive &&
            user.Username.Equals(username.Trim(), StringComparison.OrdinalIgnoreCase));

        if (account is null || !PasswordMatches(account, password))
        {
            return null;
        }

        if (!string.IsNullOrWhiteSpace(role) &&
            !account.Role.Equals(role, StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return account;
    }

    private static bool PasswordMatches(UserAccount account, string password)
    {
        return account.Password.Equals(password, StringComparison.Ordinal);
    }
}

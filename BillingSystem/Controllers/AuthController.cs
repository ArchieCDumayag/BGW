using System.Security.Claims;
using BillingSystem.Models;
using BillingSystem.Services;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace BillingSystem.Controllers;

public sealed class AuthController(IAuthService authService, IAuditLogService auditLogger) : Controller
{
    [AllowAnonymous]
    public IActionResult Login(string? returnUrl = null)
    {
        if (User.Identity?.IsAuthenticated == true)
        {
            return RedirectByRole();
        }

        return View(new LoginViewModel("", "", "Admin", returnUrl));
    }

    [HttpPost]
    [AllowAnonymous]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Login(string username, string password, string role, string? returnUrl = null)
    {
        var account = await authService.ValidateAsync(username, password, role);
        if (account is null)
        {
            await auditLogger.LogAsync(
                HttpContext,
                "Auth.LoginFailed",
                $"Failed login for username '{username?.Trim()}' with role '{role?.Trim()}'.",
                StatusCodes.Status401Unauthorized,
                username?.Trim(),
                role: role?.Trim());
            return View(new LoginViewModel(username ?? "", "", role ?? "Admin", returnUrl, "Invalid username or password."));
        }

        await SignInAsync(account);
        await auditLogger.LogAsync(
            HttpContext,
            "Auth.Login",
            "User logged in.",
            StatusCodes.Status200OK,
            account.Username,
            account.DisplayName,
            account.Role,
            account.Id);

        if (!string.IsNullOrWhiteSpace(returnUrl) && Url.IsLocalUrl(returnUrl))
        {
            return Redirect(returnUrl);
        }

        return RedirectByRole(account.Role);
    }

    [Authorize]
    public async Task<IActionResult> Logout()
    {
        await auditLogger.LogAsync(HttpContext, "Auth.Logout", "User logged out.");
        await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
        return RedirectToAction(nameof(Login));
    }

    [AllowAnonymous]
    public IActionResult Denied()
    {
        return View();
    }

    private async Task SignInAsync(UserAccount account)
    {
        var claims = new List<Claim>
        {
            new(ClaimTypes.NameIdentifier, account.Id.ToString()),
            new(ClaimTypes.Name, account.Username),
            new(ClaimTypes.Role, account.Role),
            new("DisplayName", account.DisplayName),
        };

        if (account.TechnicianId is not null)
        {
            claims.Add(new Claim("TechnicianId", account.TechnicianId.Value.ToString()));
        }

        var identity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
        await HttpContext.SignInAsync(
            CookieAuthenticationDefaults.AuthenticationScheme,
            new ClaimsPrincipal(identity));
    }

    private IActionResult RedirectByRole(string? role = null)
    {
        role ??= User.FindFirstValue(ClaimTypes.Role);
        return role?.Equals("Admin", StringComparison.OrdinalIgnoreCase) == true
            ? RedirectToAction("Dashboard", "Admin")
            : RedirectToAction("Index", "TechnicianPortal");
    }
}

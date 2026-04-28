using BillingSystem.Models;
using BillingSystem.Services;
using Microsoft.AspNetCore.Mvc;

namespace BillingSystem.Controllers;

[ApiController]
[Route("api/auth")]
public sealed class ApiAuthController(IAuthService authService) : ControllerBase
{
    [HttpPost("login")]
    public async Task<IActionResult> Login(ApiLoginRequest request)
    {
        var account = await authService.ValidateAsync(request.Username, request.Password, request.Role);
        if (account is null)
        {
            return Unauthorized(new ApiLoginResponse(false, "Invalid username or password.", "", "", "", null));
        }

        return Ok(new ApiLoginResponse(
            true,
            "Login successful.",
            account.Username,
            account.DisplayName,
            account.Role,
            EffectiveTechnicianId(account)));
    }

    private static int? EffectiveTechnicianId(UserAccount account)
    {
        if (account.TechnicianId is > 0)
        {
            return account.TechnicianId.Value;
        }

        return account.Role.Equals("Technician", StringComparison.OrdinalIgnoreCase) ? account.Id : null;
    }
}

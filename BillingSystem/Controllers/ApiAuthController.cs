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
            return Unauthorized(new ApiLoginResponse(false, "Invalid username, password, or role.", "", "", "", null));
        }

        return Ok(new ApiLoginResponse(
            true,
            "Login successful.",
            account.Username,
            account.DisplayName,
            account.Role,
            account.TechnicianId));
    }
}

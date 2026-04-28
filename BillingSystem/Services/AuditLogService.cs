using System.Security.Claims;
using BillingSystem.Models;
using Microsoft.AspNetCore.Mvc.Filters;

namespace BillingSystem.Services;

public interface IAuditLogService
{
    Task LogAsync(
        HttpContext httpContext,
        string action,
        string details = "",
        int? statusCode = null,
        string? username = null,
        string? displayName = null,
        string? role = null,
        int? userId = null);
}

public sealed class AuditLogService(IBillingStore store) : IAuditLogService
{
    private const int MaxLogEntries = 5000;

    public async Task LogAsync(
        HttpContext httpContext,
        string action,
        string details = "",
        int? statusCode = null,
        string? username = null,
        string? displayName = null,
        string? role = null,
        int? userId = null)
    {
        if (string.IsNullOrWhiteSpace(action))
        {
            return;
        }

        var user = httpContext.User;
        var routeValues = httpContext.Request.RouteValues;
        var data = await store.GetAsync();

        data.ActivityLogs.Add(new UserActivityLog
        {
            Id = NextId(data.ActivityLogs.Select(log => log.Id)),
            OccurredAt = DateTime.Now,
            UserId = userId ?? ParseNullableInt(user.FindFirstValue(ClaimTypes.NameIdentifier)),
            Username = Clean(username ?? user.FindFirstValue(ClaimTypes.Name) ?? "Anonymous"),
            DisplayName = Clean(displayName ?? user.FindFirstValue("DisplayName") ?? ""),
            Role = Clean(role ?? user.FindFirstValue(ClaimTypes.Role) ?? ""),
            Action = Clean(action),
            Controller = Clean(routeValues["controller"]?.ToString() ?? ""),
            Method = Clean(httpContext.Request.Method),
            Path = Clean(httpContext.Request.Path.ToString()),
            IpAddress = Clean(httpContext.Connection.RemoteIpAddress?.ToString() ?? ""),
            StatusCode = statusCode ?? httpContext.Response.StatusCode,
            Details = Clean(details)
        });

        if (data.ActivityLogs.Count > MaxLogEntries)
        {
            data.ActivityLogs = data.ActivityLogs
                .OrderByDescending(log => log.OccurredAt)
                .ThenByDescending(log => log.Id)
                .Take(MaxLogEntries)
                .OrderBy(log => log.OccurredAt)
                .ThenBy(log => log.Id)
                .ToList();
        }

        await store.SaveAsync(data);
    }

    private static int NextId(IEnumerable<int> ids)
    {
        return ids.DefaultIfEmpty().Max() + 1;
    }

    private static int? ParseNullableInt(string? value)
    {
        return int.TryParse(value, out var number) ? number : null;
    }

    private static string Clean(string value)
    {
        return value.Trim();
    }
}

public sealed class AuditLogActionFilter(IAuditLogService auditLogger) : IAsyncActionFilter
{
    public async Task OnActionExecutionAsync(ActionExecutingContext context, ActionExecutionDelegate next)
    {
        var controller = context.RouteData.Values["controller"]?.ToString() ?? "";
        var action = context.RouteData.Values["action"]?.ToString() ?? "";
        var shouldLog =
            context.HttpContext.User.Identity?.IsAuthenticated == true &&
            !controller.Equals("Auth", StringComparison.OrdinalIgnoreCase);

        var executedContext = await next();

        if (!shouldLog)
        {
            return;
        }

        var statusCode = executedContext.Exception is not null && !executedContext.ExceptionHandled
            ? 500
            : context.HttpContext.Response.StatusCode;

        await auditLogger.LogAsync(
            context.HttpContext,
            $"{controller}.{action}",
            $"{context.HttpContext.Request.Method} {controller}/{action}",
            statusCode);
    }
}

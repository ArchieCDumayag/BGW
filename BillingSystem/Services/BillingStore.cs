using System.Text.Json;
using BillingSystem.Models;

namespace BillingSystem.Services;

public interface IBillingStore
{
    Task<BillingData> GetAsync();
    Task SaveAsync(BillingData data);
    Task<Client?> GetClientAsync(int id);
    Task<JobTicket?> GetJobAsync(int id);
}

public sealed class JsonBillingStore(IWebHostEnvironment environment) : IBillingStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    private readonly string _path = Path.Combine(environment.ContentRootPath, "Data", "billing-data.json");
    private readonly SemaphoreSlim _gate = new(1, 1);
    private BillingData? _cache;

    public async Task<BillingData> GetAsync()
    {
        await _gate.WaitAsync();
        try
        {
            if (_cache is not null)
            {
                return _cache;
            }

            if (!File.Exists(_path))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_path)!);
                _cache = new BillingData();
                await SaveUnlockedAsync(_cache);
                return _cache;
            }

            await using var stream = File.OpenRead(_path);
            _cache = await JsonSerializer.DeserializeAsync<BillingData>(stream, JsonOptions) ?? new BillingData();
            Normalize(_cache);
            return _cache;
        }
        finally
        {
            _gate.Release();
        }
    }

    public async Task SaveAsync(BillingData data)
    {
        await _gate.WaitAsync();
        try
        {
            Normalize(data);
            _cache = data;
            await SaveUnlockedAsync(data);
        }
        finally
        {
            _gate.Release();
        }
    }

    public async Task<Client?> GetClientAsync(int id)
    {
        var data = await GetAsync();
        return data.Clients.FirstOrDefault(client => client.Id == id);
    }

    public async Task<JobTicket?> GetJobAsync(int id)
    {
        var data = await GetAsync();
        return data.Jobs.FirstOrDefault(job => job.Id == id);
    }

    private async Task SaveUnlockedAsync(BillingData data)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(_path)!);
        var tempPath = $"{_path}.tmp";
        await using (var stream = File.Create(tempPath))
        {
            await JsonSerializer.SerializeAsync(stream, data, JsonOptions);
        }

        File.Copy(tempPath, _path, overwrite: true);
        File.Delete(tempPath);
    }

    private static void Normalize(BillingData data)
    {
        data.Clients ??= [];
        data.Payments ??= [];
        data.Expenses ??= [];
        data.Technicians ??= [];
        data.Jobs ??= [];
        data.PppoeUsers ??= [];
        data.TrafficSamples ??= [];
        data.UserAccounts ??= [];
        data.PlanChanges ??= [];
        data.MonthlyBillOverrides ??= [];
        data.Referrals ??= [];
        data.Plans ??= [];
        data.CoverageAreas ??= [];
        data.NapLocations ??= [];
        data.OltDevices ??= [];
        data.OltPonPorts ??= [];
        data.OltOnuClients ??= [];
        data.CollectorAssignments ??= [];
        data.Tickets ??= [];
        data.ActivityLogs ??= [];
        data.Settings ??= new SystemSettings();

        foreach (var client in data.Clients)
        {
            NormalizeClient(client);
        }

        foreach (var payment in data.Payments)
        {
            payment.Method ??= "";
            payment.ReferenceNumber ??= "";
            payment.CollectedBy ??= "";
            payment.Remarks ??= "";
        }

        foreach (var overrideBill in data.MonthlyBillOverrides)
        {
            overrideBill.DiscountRemarks ??= "";
            overrideBill.Remarks ??= "";
        }

        foreach (var referral in data.Referrals)
        {
            referral.ReferrerName ??= "";
            referral.NewClientName ??= "";
            referral.ReferralText ??= "";
            referral.Remarks ??= "";
        }

        foreach (var plan in data.Plans)
        {
            plan.PlanName ??= "";
            plan.Type = string.IsNullOrWhiteSpace(plan.Type) ? "Prepaid" : plan.Type;
        }

        foreach (var area in data.CoverageAreas)
        {
            area.AreaName ??= "";
        }

        foreach (var nap in data.NapLocations)
        {
            nap.Name ??= "";
            nap.AreaName ??= "";
            nap.Remarks ??= "";
        }

        foreach (var olt in data.OltDevices)
        {
            olt.OltName ??= "";
            olt.Technology = string.IsNullOrWhiteSpace(olt.Technology) ? "Gpon" : olt.Technology;
            olt.Site ??= "";
            olt.ManagementUrl ??= "";
            olt.Username ??= "";
            olt.Password ??= "";
            olt.TotalPonPorts = Math.Max(0, olt.TotalPonPorts);
        }

        var oltTechnologyById = data.OltDevices.ToDictionary(olt => olt.Id, olt => olt.Technology);
        foreach (var ponPort in data.OltPonPorts)
        {
            ponPort.PonPort ??= "";
            ponPort.CustomerCapacity = oltTechnologyById.TryGetValue(ponPort.OltDeviceId, out var technology)
                ? CustomerCapacityForTechnology(technology)
                : Math.Max(0, ponPort.CustomerCapacity);
            ponPort.TotalNap = Math.Max(0, ponPort.TotalNap);
        }

        foreach (var oltClient in data.OltOnuClients)
        {
            oltClient.OltName ??= "";
            oltClient.PonPort ??= "";
            oltClient.OnuId ??= "";
            oltClient.Status ??= "";
            oltClient.Description ??= "";
            oltClient.Model ??= "";
            oltClient.Profile ??= "";
            oltClient.AuthMode ??= "";
            oltClient.SerialNumber ??= "";
        }

        foreach (var assignment in data.CollectorAssignments)
        {
            assignment.CollectorName ??= "";
            assignment.AreaName ??= "";
            assignment.Remarks ??= "";
        }

        foreach (var ticket in data.Tickets)
        {
            ticket.Subject ??= "";
            ticket.Type = string.IsNullOrWhiteSpace(ticket.Type) ? "Repair" : ticket.Type;
            ticket.Priority = string.IsNullOrWhiteSpace(ticket.Priority) ? "Normal" : ticket.Priority;
            ticket.Status = string.IsNullOrWhiteSpace(ticket.Status) ? "Open" : ticket.Status;
            ticket.AssignedTo ??= "";
            ticket.Remarks ??= "";
        }

        foreach (var log in data.ActivityLogs)
        {
            log.Username ??= "";
            log.DisplayName ??= "";
            log.Role ??= "";
            log.Action ??= "";
            log.Controller ??= "";
            log.Method ??= "";
            log.Path ??= "";
            log.IpAddress ??= "";
            log.Details ??= "";
        }

        data.Settings.SemaphoreApiKey ??= "";
        data.Settings.SemaphoreSenderName ??= "";
        data.Settings.CompanyName = string.IsNullOrWhiteSpace(data.Settings.CompanyName) ? "Billing System" : data.Settings.CompanyName;
        data.Settings.SystemDisplayName = string.IsNullOrWhiteSpace(data.Settings.SystemDisplayName)
            ? data.Settings.CompanyName
            : data.Settings.SystemDisplayName;
        data.Settings.CompanyLogoUrl ??= "";
        data.Settings.BrowserLogoUrl ??= "";
        data.Settings.MikrotikHost ??= "";
        if (data.Settings.MikrotikApiPort <= 0)
        {
            data.Settings.MikrotikApiPort = 8728;
        }

        data.Settings.MikrotikApiUser ??= "";
        data.Settings.MikrotikApiPassword ??= "";
        data.Settings.GCashAccountName ??= "";
        data.Settings.GCashAccountNumber ??= "";
        data.Settings.GCashQrCodeUrl ??= "";
    }

    private static void NormalizeClient(Client client)
    {
        client.AccountNumber ??= "";
        client.Status ??= "Active";
        client.BillingType = BillingRules.NormalizeBillingType(client.BillingType);
        client.Area ??= "";
        client.Zone ??= "";
        client.Name ??= "";
        client.PppoeUsername ??= "";
        client.Contact ??= "";
        client.Email ??= "";
        client.FacebookAccount ??= "";
        client.Referral = string.IsNullOrWhiteSpace(client.Referral) ? "INQUIRE" : client.Referral;
        client.Address ??= "";
        client.Remarks ??= "";
    }

    private static int CustomerCapacityForTechnology(string? technology)
    {
        return (technology ?? "").Contains("epon", StringComparison.OrdinalIgnoreCase) ? 32 : 128;
    }
}

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
        data.CollectorAssignments ??= [];
        data.Tickets ??= [];
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

        data.Settings.SemaphoreApiKey ??= "";
        data.Settings.SemaphoreSenderName ??= "";
        data.Settings.MikrotikHost ??= "";
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
        client.BillingType ??= "Prepaid";
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
}

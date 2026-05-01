using System.Globalization;
using System.Security.Claims;
using System.Text.Json;
using System.Text.RegularExpressions;
using BillingSystem.Models;
using BillingSystem.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace BillingSystem.Controllers;

[Authorize(Roles = "Admin")]
public sealed class AdminController(
    IBillingStore store,
    IWebHostEnvironment environment,
    IAuditLogService auditLogger,
    IOltWebClient oltWebClient) : Controller
{
    private const string ClearClientsConfirmation = "CLEAR ALL";
    private const string StatementCompanyName = "3J COMPUTER AND INTERNET INSTALLATION SERVICES";
    private const string StatementCompanyAddress = "Zone 7, Poblacion, Baggao, Cagayan";
    private const string StatementCompanyContact = "0965-140-4623 FR****S VI*L D. / 0936-156-5251 AR***E D.";
    private const string ThermalSupportContact = "09651404623 / 09361565251";
    private static readonly string[] AccountRoles = ["Admin", "User", "Technician", "Collector"];

    public async Task<IActionResult> Dashboard()
    {
        var data = await store.GetAsync();
        var today = DateOnly.FromDateTime(DateTime.Today);
        var monthPayments = data.Payments
            .Where(p => p.PaidOn.Year == today.Year && p.PaidOn.Month == today.Month)
            .ToList();
        var totalBilled = data.Clients.Sum(c => c.Bills);
        var totalCollected = monthPayments.Sum(p => p.Amount);
        var methodSummaries = BuildMethodSummaries(monthPayments, totalCollected);
        var clientsById = data.Clients.ToDictionary(c => c.Id);

        var model = new DashboardViewModel
        {
            TotalClients = data.Clients.Count,
            ActiveClients = data.Clients.Count(c => c.Status.Equals("Active", StringComparison.OrdinalIgnoreCase)),
            DisconnectedClients = data.Clients.Count(c => c.Status.Equals("DC", StringComparison.OrdinalIgnoreCase) || c.Status.Contains("disconnect", StringComparison.OrdinalIgnoreCase)),
            MonthlyRecurringRevenue = data.Clients.Where(c => c.Status.Equals("Active", StringComparison.OrdinalIgnoreCase)).Sum(c => c.PlanAmount),
            TotalBalances = data.Clients.Sum(c => c.Balance),
            PaymentsThisMonth = totalCollected,
            ExpensesThisMonth = data.Expenses.Where(e => e.SpentOn.Year == today.Year && e.SpentOn.Month == today.Month).Sum(e => e.Amount),
            OpenJobs = data.Jobs.Count(j => !j.Status.Equals("Done", StringComparison.OrdinalIgnoreCase)),
            CurrentMonthLabel = new DateTime(today.Year, today.Month, 1).ToString("MMMM yyyy"),
            TotalBilled = totalBilled,
            TotalCollected = totalCollected,
            CollectionRate = totalBilled == 0 ? 0 : totalCollected / totalBilled * 100,
            PayingSubscribers = monthPayments.Where(p => p.ClientId > 0).Select(p => p.ClientId).Distinct().Count(),
            OutstandingBalance = data.Clients.Sum(c => c.Balance),
            CashCollected = methodSummaries.First(m => m.Method == "Cash").Amount,
            GCashCollected = methodSummaries.First(m => m.Method == "GCash").Amount,
            OtherCollected = methodSummaries.First(m => m.Method == "Other").Amount,
            MethodSummaries = methodSummaries,
            RecentPayments = monthPayments
                .OrderByDescending(p => p.PaidOn)
                .ThenByDescending(p => p.Id)
                .Take(10)
                .Select(p => new DashboardPaymentEntry
                {
                    PaidOn = p.PaidOn,
                    ClientName = clientsById.TryGetValue(p.ClientId, out var client) ? client.Name : "Unmatched client",
                    Method = NormalizePaymentMethod(p.Method),
                    Amount = p.Amount,
                    ReferenceNumber = p.ReferenceNumber
                })
                .ToList(),
            RecentClients = data.Clients.OrderByDescending(c => c.Id).Take(8).ToList(),
            RecentJobs = data.Jobs.OrderByDescending(j => j.Id).Take(8).ToList()
        };

        return View(model);
    }

    public async Task<IActionResult> Clients(string? q, string? status, string? area, string? type, string? sort)
    {
        var data = await store.GetAsync();
        var clients = data.Clients.AsEnumerable();

        if (!string.IsNullOrWhiteSpace(q))
        {
            clients = clients.Where(c => (c.Name ?? "").Contains(q, StringComparison.OrdinalIgnoreCase)
                || (c.PppoeUsername ?? "").Contains(q, StringComparison.OrdinalIgnoreCase)
                || (c.AccountNumber ?? "").Contains(q, StringComparison.OrdinalIgnoreCase));
        }

        if (!string.IsNullOrWhiteSpace(status))
        {
            clients = clients.Where(c => (c.Status ?? "").Equals(status, StringComparison.OrdinalIgnoreCase));
        }

        if (!string.IsNullOrWhiteSpace(area))
        {
            clients = clients.Where(c => (c.Area ?? "").Equals(area, StringComparison.OrdinalIgnoreCase));
        }

        if (!string.IsNullOrWhiteSpace(type))
        {
            clients = clients.Where(c => (c.BillingType ?? "").Equals(type, StringComparison.OrdinalIgnoreCase));
        }

        ViewBag.Query = q ?? "";
        ViewBag.Status = status ?? "";
        ViewBag.Area = area ?? "";
        ViewBag.Type = type ?? "";
        ViewBag.Sort = sort ?? "";
        ViewBag.Areas = data.Clients.Select(c => c.Area).Where(a => !string.IsNullOrWhiteSpace(a)).Distinct().Order().ToList();
        ViewBag.Types = data.Clients.Select(c => c.BillingType).Where(t => !string.IsNullOrWhiteSpace(t)).Distinct().Order().ToList();
        ViewBag.ClientBillingRules = data.Clients.ToDictionary(c => c.Id, c => BillingRules.ForClient(c));
        ViewBag.ClientDeleteCaptchas = data.Clients.ToDictionary(c => c.Id, DeleteCaptchaQuestion);
        ViewBag.NextAccountNumber = NextAccountNumber(data.Clients);
        ViewBag.ReferralClients = data.Clients.OrderBy(c => c.Name ?? "").ThenBy(c => AccountSortKey(c.AccountNumber)).ToList();

        clients = sort switch
        {
            "account" => clients.OrderBy(c => AccountSortKey(c.AccountNumber)).ThenBy(c => c.AccountNumber ?? "").ThenBy(c => c.Name ?? ""),
            "account-desc" => clients.OrderByDescending(c => AccountSortKey(c.AccountNumber)).ThenByDescending(c => c.AccountNumber ?? "").ThenBy(c => c.Name ?? ""),
            "name" => clients.OrderBy(c => c.Name ?? ""),
            "name-desc" => clients.OrderByDescending(c => c.Name ?? ""),
            "type" => clients.OrderBy(c => c.BillingType ?? "").ThenBy(c => c.Name ?? ""),
            "type-desc" => clients.OrderByDescending(c => c.BillingType ?? "").ThenBy(c => c.Name ?? ""),
            "status" => clients.OrderBy(c => c.Status ?? "").ThenBy(c => c.Name ?? ""),
            "status-desc" => clients.OrderByDescending(c => c.Status ?? "").ThenBy(c => c.Name ?? ""),
            "plan" => clients.OrderBy(c => c.PlanAmount).ThenBy(c => c.Name ?? ""),
            "plan-desc" => clients.OrderByDescending(c => c.PlanAmount).ThenBy(c => c.Name ?? ""),
            "area" => clients.OrderBy(c => c.Area ?? "").ThenBy(c => c.Zone ?? "").ThenBy(c => c.Name ?? ""),
            "area-desc" => clients.OrderByDescending(c => c.Area ?? "").ThenBy(c => c.Zone ?? "").ThenBy(c => c.Name ?? ""),
            "pppoe" => clients.OrderBy(c => string.IsNullOrWhiteSpace(c.PppoeUsername)).ThenBy(c => c.PppoeUsername ?? "").ThenBy(c => c.Name ?? ""),
            "pppoe-desc" => clients.OrderBy(c => string.IsNullOrWhiteSpace(c.PppoeUsername)).ThenByDescending(c => c.PppoeUsername ?? "").ThenBy(c => c.Name ?? ""),
            _ => clients.OrderBy(c => AccountSortKey(c.AccountNumber)).ThenBy(c => c.AccountNumber ?? "").ThenBy(c => c.Name ?? "")
        };

        return View(clients.ToList());
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ImportClientList(IFormFile clientListFile, int importYear)
    {
        if (clientListFile is null || clientListFile.Length == 0)
        {
            TempData["ClientError"] = "Please choose a client list Excel file to import.";
            return RedirectToAction(nameof(Clients));
        }

        if (!Path.GetExtension(clientListFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            TempData["ClientError"] = "Client list import only accepts .xlsx files.";
            return RedirectToAction(nameof(Clients));
        }

        if (importYear < 2000 || importYear > 2100)
        {
            TempData["ClientError"] = "Import year must be between 2000 and 2100.";
            return RedirectToAction(nameof(Clients));
        }

        var importsPath = Path.Combine(environment.ContentRootPath, "Data", "imports");
        Directory.CreateDirectory(importsPath);
        var savedFileName = $"client-list-{importYear}-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx";
        var savedPath = Path.Combine(importsPath, savedFileName);

        await using (var uploadStream = System.IO.File.Create(savedPath))
        {
            await clientListFile.CopyToAsync(uploadStream);
        }

        try
        {
            var data = await store.GetAsync();
            await using var readStream = System.IO.File.OpenRead(savedPath);
            BackupBillingDataFile("client-import");
            var result = ClientListExcelImporter.ReplaceClients(data, readStream, importYear, savedFileName);
            ClearClientOperationalRecords(data);
            await store.SaveAsync(data);

            TempData["ClientMessage"] =
                $"Replaced client list from {savedFileName}: {result.Imported:N0} clients imported. " +
                $"Imported {result.PaymentsImported:N0} payments and {result.MonthlyBillsImported:N0} monthly bills. " +
                $"Referral details recorded: {result.ReferralsImported:N0}. " +
                $"Unmatched payments: {result.UnmatchedPayments:N0}. " +
                $"Skipped {result.SkippedDuplicate:N0} duplicate and {result.SkippedInvalid:N0} invalid rows.";
        }
        catch (Exception ex)
        {
            TempData["ClientError"] = $"Import failed: {ex.Message}";
        }

        return RedirectToAction(nameof(Clients));
    }

    private void BackupBillingDataFile(string reason)
    {
        var dataPath = BillingDataPath();
        if (!System.IO.File.Exists(dataPath))
        {
            return;
        }

        var backupPath = Path.Combine(
            environment.ContentRootPath,
            "Data",
            $"billing-data.{reason}-backup-{DateTime.Now:yyyyMMdd-HHmmss-fff}.json");
        System.IO.File.Copy(dataPath, backupPath);
    }

    private string BillingDataPath()
    {
        return Path.Combine(environment.ContentRootPath, "Data", "billing-data.json");
    }

    private async Task<SystemSettings> MergeSettingsAsync(
        SystemSettings current,
        SystemSettings posted,
        IFormFile? companyLogoFile,
        IFormFile? browserLogoFile)
    {
        posted.CompanyName = string.IsNullOrWhiteSpace(posted.CompanyName) ? "Billing System" : posted.CompanyName.Trim();
        posted.SystemDisplayName = string.IsNullOrWhiteSpace(posted.SystemDisplayName)
            ? posted.CompanyName
            : posted.SystemDisplayName.Trim();
        posted.CompanyLogoUrl = string.IsNullOrWhiteSpace(posted.CompanyLogoUrl) ? current.CompanyLogoUrl : posted.CompanyLogoUrl.Trim();
        posted.BrowserLogoUrl = string.IsNullOrWhiteSpace(posted.BrowserLogoUrl) ? current.BrowserLogoUrl : posted.BrowserLogoUrl.Trim();
        posted.SmsReminderTemplate = posted.SmsReminderTemplate?.Trim() ?? "";
        posted.Currency = string.IsNullOrWhiteSpace(posted.Currency) ? "PHP" : posted.Currency.Trim();
        posted.SemaphoreApiKey = posted.SemaphoreApiKey?.Trim() ?? "";
        posted.SemaphoreSenderName = posted.SemaphoreSenderName?.Trim() ?? "";
        posted.MikrotikHost = posted.MikrotikHost?.Trim() ?? "";
        posted.MikrotikApiUser = posted.MikrotikApiUser?.Trim() ?? "";
        posted.MikrotikApiPassword = string.IsNullOrWhiteSpace(posted.MikrotikApiPassword)
            ? current.MikrotikApiPassword
            : posted.MikrotikApiPassword;
        posted.GCashAccountName = posted.GCashAccountName?.Trim() ?? "";
        posted.GCashAccountNumber = posted.GCashAccountNumber?.Trim() ?? "";
        posted.GCashQrCodeUrl = posted.GCashQrCodeUrl?.Trim() ?? "";

        if (companyLogoFile is { Length: > 0 })
        {
            posted.CompanyLogoUrl = await SaveBrandingFileAsync(companyLogoFile, "company-logo");
        }

        if (browserLogoFile is { Length: > 0 })
        {
            posted.BrowserLogoUrl = await SaveBrandingFileAsync(browserLogoFile, "browser-logo");
        }

        return posted;
    }

    private async Task<string> SaveBrandingFileAsync(IFormFile file, string name)
    {
        var extension = Path.GetExtension(file.FileName).ToLowerInvariant();
        var allowedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".png",
            ".jpg",
            ".jpeg",
            ".gif",
            ".webp",
            ".ico",
            ".svg"
        };

        if (!allowedExtensions.Contains(extension))
        {
            throw new InvalidOperationException("Branding files must be image files: png, jpg, jpeg, gif, webp, ico, or svg.");
        }

        if (file.Length > 2 * 1024 * 1024)
        {
            throw new InvalidOperationException("Branding image must be 2 MB or smaller.");
        }

        var uploadRoot = Path.Combine(environment.WebRootPath, "uploads", "branding");
        Directory.CreateDirectory(uploadRoot);
        var fileName = $"{name}-{DateTime.Now:yyyyMMddHHmmssfff}{extension}";
        var filePath = Path.Combine(uploadRoot, fileName);

        await using var stream = System.IO.File.Create(filePath);
        await file.CopyToAsync(stream);

        return $"/uploads/branding/{fileName}";
    }

    private static void ClearClientOperationalRecords(BillingData data)
    {
        data.Jobs.Clear();
        data.PppoeUsers.Clear();
        data.TrafficSamples.Clear();
    }

    public async Task<IActionResult> ClientHistory(int id)
    {
        var data = await store.GetAsync();
        var client = data.Clients.FirstOrDefault(c => c.Id == id);
        if (client is null)
        {
            return NotFound();
        }

        var payments = data.Payments
            .Where(p => p.ClientId == id)
            .OrderByDescending(p => p.PaidOn)
            .ThenByDescending(p => p.Id)
            .ToList();

        var model = new ClientPaymentHistoryViewModel
        {
            Client = client,
            Pppoe = data.PppoeUsers.FirstOrDefault(p => p.ClientId == id) ?? new PppoeUser
            {
                ClientId = id,
                Username = client.PppoeUsername,
                Status = client.Status
            },
            BillingRule = BillingRules.ForClient(client),
            CollectionStatus = BillingRules.CollectionStatusForClient(client, payments),
            Payments = payments,
            TotalPaid = payments.Sum(p => p.Amount)
        };

        return View(model);
    }

    public async Task<IActionResult> CustomerAccount(int id, int? year)
    {
        var data = await store.GetAsync();
        var client = data.Clients.FirstOrDefault(c => c.Id == id);
        if (client is null)
        {
            return NotFound();
        }

        var payments = data.Payments
            .Where(p => p.ClientId == id)
            .OrderByDescending(p => p.PaidOn)
            .ThenByDescending(p => p.Id)
            .ToList();
        var planChanges = data.PlanChanges
            .Where(change => change.ClientId == id)
            .OrderByDescending(change => change.EffectiveMonth)
            .ThenByDescending(change => change.Id)
            .ToList();
        var monthlyBillOverrides = data.MonthlyBillOverrides
            .Where(overrideBill => overrideBill.ClientId == id)
            .OrderByDescending(overrideBill => overrideBill.BillingMonth)
            .ThenByDescending(overrideBill => overrideBill.Id)
            .ToList();

        var allBillingMonths = BuildCustomerBillingMonths(client, payments, planChanges, monthlyBillOverrides);
        var availableYears = allBillingMonths
            .Select(month => month.Month.Year)
            .Distinct()
            .OrderByDescending(value => value)
            .ToList();
        var today = DateOnly.FromDateTime(DateTime.Today);
        var selectedYear = year is int requestedYear && availableYears.Contains(requestedYear)
            ? requestedYear
            : today.Year;

        if (!availableYears.Contains(selectedYear) && availableYears.Count > 0)
        {
            selectedYear = availableYears[0];
        }

        var model = new CustomerAccountViewModel
        {
            Client = client,
            BillingRule = BillingRules.ForClient(client),
            CurrentBalance = client.Balance,
            PlanChanges = planChanges,
            MonthlyBillOverrides = monthlyBillOverrides,
            BillingMonths = allBillingMonths.Where(month => month.Month.Year == selectedYear).ToList(),
            PaymentHistory = payments.Where(payment => payment.PaidOn.Year == selectedYear).ToList(),
            SelectedYear = selectedYear,
            AvailableYears = availableYears
        };

        return View(model);
    }

    public async Task<IActionResult> AccountStatement(int id, int? year)
    {
        var data = await store.GetAsync();
        var client = data.Clients.FirstOrDefault(c => c.Id == id);
        if (client is null)
        {
            return NotFound();
        }

        var model = BuildCustomerStatementModel(data, client, "Account Statement", year, includeFullYear: true);
        return View(model);
    }

    public async Task<IActionResult> BillingStatement(int id)
    {
        var data = await store.GetAsync();
        var client = data.Clients.FirstOrDefault(c => c.Id == id);
        if (client is null)
        {
            return NotFound();
        }

        var model = BuildCustomerStatementModel(data, client, "Billing Statement", null, includeFullYear: false);
        return View(model);
    }

    public async Task<IActionResult> ThermalReceipt(int id, int? paymentId, bool autoPrint = false)
    {
        var data = await store.GetAsync();
        var client = data.Clients.FirstOrDefault(c => c.Id == id);
        if (client is null)
        {
            return NotFound();
        }

        if (paymentId.HasValue && data.Payments.All(payment => payment.Id != paymentId.Value || payment.ClientId != client.Id))
        {
            return NotFound();
        }

        ViewBag.AutoPrint = autoPrint;
        var model = BuildThermalReceiptModel(data, client, paymentId);
        return View(model);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ChangeClientPlan(int id, decimal planAmount, string? effectiveMonth)
    {
        var data = await store.GetAsync();
        var client = data.Clients.FirstOrDefault(c => c.Id == id);
        if (client is null)
        {
            return NotFound();
        }

        if (planAmount <= 0)
        {
            TempData["CustomerAccountError"] = "Plan amount must be greater than zero.";
            return RedirectToAction(nameof(CustomerAccount), new { id });
        }

        var today = DateOnly.FromDateTime(DateTime.Today);
        var currentMonth = new DateOnly(today.Year, today.Month, 1);
        var effective = ParseMonth(effectiveMonth) ?? currentMonth;
        if (effective < currentMonth)
        {
            effective = currentMonth;
        }

        var clientPayments = data.Payments.Where(p => p.ClientId == id).ToList();
        EnsurePlanBaseline(data, client, clientPayments);
        UpsertPlanChange(data, client.Id, effective, planAmount, "Plan changed from customer account.");

        client.PlanAmount = PlanAmountForBillingMonth(client, currentMonth, data.PlanChanges.Where(change => change.ClientId == id));

        await store.SaveAsync(data);
        TempData["CustomerAccountMessage"] = $"Plan updated to PHP {planAmount:N0} starting {effective:MMMM yyyy}. Previous months keep their original bill amount.";
        return RedirectToAction(nameof(CustomerAccount), new { id });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> UpdateMonthlyBill(
        int id,
        string? billingMonth,
        decimal billAmount,
        int? returnYear,
        decimal? discountAmount,
        string? discountRemarks)
    {
        var data = await store.GetAsync();
        var client = data.Clients.FirstOrDefault(c => c.Id == id);
        if (client is null)
        {
            return NotFound();
        }

        var month = ParseMonth(billingMonth);
        if (month is null)
        {
            TempData["CustomerAccountError"] = "Billing month is invalid.";
            return RedirectToAction(nameof(CustomerAccount), new { id });
        }

        if (billAmount < 0)
        {
            TempData["CustomerAccountError"] = "Current bill cannot be negative.";
            return RedirectToAction(nameof(CustomerAccount), new { id });
        }

        decimal? appliedDiscount = null;
        string? appliedDiscountRemarks = null;
        if (discountAmount.HasValue)
        {
            appliedDiscount = Math.Max(0, discountAmount.Value);
            appliedDiscountRemarks = appliedDiscount.Value > 0
                ? string.IsNullOrWhiteSpace(discountRemarks)
                    ? $"Admin approved PHP {appliedDiscount:N0} discount."
                    : discountRemarks.Trim()
                : "";
        }

        UpsertMonthlyBillOverride(
            data,
            client.Id,
            month.Value,
            billAmount,
            "Current bill changed from customer account.",
            appliedDiscount,
            appliedDiscountRemarks);

        await store.SaveAsync(data);
        TempData["CustomerAccountMessage"] = $"Current bill for {month.Value:MMMM yyyy} updated to PHP {billAmount:N0}.";
        return RedirectToAction(nameof(CustomerAccount), new { id, year = returnYear ?? month.Value.Year });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SaveClient(Client client)
    {
        var data = await store.GetAsync();
        if (client.Id == 0)
        {
            client.Id = NextId(data.Clients.Select(c => c.Id));
            client.AccountNumber = NextAccountNumber(data.Clients);
            client.DateInstalled ??= DateOnly.FromDateTime(DateTime.Today);
            client.Status = string.IsNullOrWhiteSpace(client.Status) ? "Active" : client.Status;
            client.BillingType = BillingRules.NormalizeBillingType(client.BillingType);
            client.Referral = ReferralBillingService.NormalizeReferralText(client.Referral);
            client.Bills = BillingRules.ProratedFirstBill(client.PlanAmount, client.DateInstalled.Value, client.BillingType);
            if (client.Balance <= 0)
            {
                client.Balance = Math.Max(0, client.Bills - client.Advance);
            }

            data.Clients.Add(client);
            UpsertMonthlyBillOverride(
                data,
                client.Id,
                new DateOnly(client.DateInstalled.Value.Year, client.DateInstalled.Value.Month, 1),
                client.Bills,
                $"{client.BillingType} prorated first bill from installation date {client.DateInstalled.Value:MMM dd, yyyy}.");

            var referralResult = ReferralBillingService.ApplyReferralDiscount(data, client);
            if (!string.IsNullOrWhiteSpace(referralResult))
            {
                TempData["ClientMessage"] = referralResult;
            }
        }
        else
        {
            var existing = data.Clients.FirstOrDefault(c => c.Id == client.Id);
            if (existing is null)
            {
                return NotFound();
            }

            existing.AccountNumber = client.AccountNumber;
            existing.DateInstalled = client.DateInstalled;
            existing.Name = client.Name;
            existing.Status = client.Status;
            existing.BillingType = BillingRules.NormalizeBillingType(client.BillingType);
            if (existing.PlanAmount != client.PlanAmount)
            {
                var today = DateOnly.FromDateTime(DateTime.Today);
                var currentMonth = new DateOnly(today.Year, today.Month, 1);
                var clientPayments = data.Payments.Where(p => p.ClientId == existing.Id).ToList();
                EnsurePlanBaseline(data, existing, clientPayments);
                UpsertPlanChange(data, existing.Id, currentMonth, client.PlanAmount, "Plan changed from clients list.");
            }

            existing.PlanAmount = client.PlanAmount;
            existing.Area = client.Area;
            existing.Zone = client.Zone;
            existing.PppoeUsername = client.PppoeUsername;
            existing.Contact = client.Contact;
            existing.Address = client.Address;
            existing.Balance = client.Balance;
            existing.Advance = client.Advance;
            existing.Bills = client.Bills;
            existing.Remarks = client.Remarks;
        }

        await store.SaveAsync(data);
        return RedirectToAction(nameof(Clients), new { q = client.Name });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteClient(int id, string? captchaAnswer)
    {
        var data = await store.GetAsync();
        var client = data.Clients.FirstOrDefault(c => c.Id == id);
        if (client is null)
        {
            return NotFound();
        }

        if (!string.Equals(captchaAnswer?.Trim(), DeleteCaptchaAnswer(client), StringComparison.OrdinalIgnoreCase))
        {
            TempData["ClientError"] = $"Wrong captcha answer. {client.Name} was not deleted.";
            return RedirectToAction(nameof(Clients), new { q = client.Name });
        }

        BackupBillingDataFile("client-delete");
        data.Clients.Remove(client);
        data.Payments.RemoveAll(p => p.ClientId == id);
        data.PlanChanges.RemoveAll(change => change.ClientId == id);
        data.MonthlyBillOverrides.RemoveAll(overrideBill => overrideBill.ClientId == id);
        data.Referrals.RemoveAll(referral => referral.ReferrerClientId == id || referral.NewClientId == id);
        data.PppoeUsers.RemoveAll(p => p.ClientId == id);
        data.TrafficSamples.RemoveAll(t => t.ClientId == id);
        data.Jobs.RemoveAll(j => j.ClientId == id);

        await store.SaveAsync(data);
        TempData["ClientMessage"] = $"{client.Name} and related records were deleted.";
        return RedirectToAction(nameof(Clients));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ClearClientRecords(string? confirmation)
    {
        if (!string.Equals(confirmation?.Trim(), ClearClientsConfirmation, StringComparison.OrdinalIgnoreCase))
        {
            TempData["ClientError"] = $"Type {ClearClientsConfirmation} to clear all client records.";
            return RedirectToAction(nameof(Clients));
        }

        var data = await store.GetAsync();
        var clientCount = data.Clients.Count;
        var paymentCount = data.Payments.Count;
        var monthlyBillCount = data.MonthlyBillOverrides.Count;

        BackupBillingDataFile("clients-clear");
        ClearAllClientRecords(data);

        await store.SaveAsync(data);
        TempData["ClientMessage"] =
            $"Cleared {clientCount:N0} clients, {paymentCount:N0} payments, and {monthlyBillCount:N0} monthly bills. A backup was created before clearing.";
        return RedirectToAction(nameof(Clients));
    }

    private static void ClearAllClientRecords(BillingData data)
    {
        data.Clients.Clear();
        data.Payments.Clear();
        data.PlanChanges.Clear();
        data.MonthlyBillOverrides.Clear();
        data.Referrals.Clear();
        data.Jobs.Clear();
        data.PppoeUsers.Clear();
        data.TrafficSamples.Clear();
    }

    public async Task<IActionResult> Payments(string? month)
    {
        var data = await store.GetAsync();
        var today = DateOnly.FromDateTime(DateTime.Today);
        var currentMonth = new DateOnly(today.Year, today.Month, 1);
        var selectedMonth = ParseMonth(month) ?? currentMonth;
        ViewBag.Clients = data.Clients.OrderBy(c => c.Name).ToList();
        ViewBag.ClientMap = data.Clients.ToDictionary(c => c.Id);
        ViewBag.ClientBillingRules = data.Clients.ToDictionary(c => c.Id, c => BillingRules.ForClient(c));
        ViewBag.ClientCurrentBills = CurrentBillAmountsForMonth(data, selectedMonth);
        ViewBag.SelectedMonthValue = selectedMonth.ToString("yyyy-MM", CultureInfo.InvariantCulture);
        ViewBag.SelectedMonthLabel = selectedMonth.ToString("MMMM yyyy", CultureInfo.InvariantCulture);
        return View(BuildClientCurrentBillRows(data, selectedMonth));
    }

    public async Task<IActionResult> PaymentHistory()
    {
        var data = await store.GetAsync();
        ViewBag.ClientMap = data.Clients.ToDictionary(c => c.Id);
        ViewBag.PaymentTableTitle = "Payment history";
        return View(data.Payments.OrderByDescending(p => p.PaidOn).ThenByDescending(p => p.Id).ToList());
    }

    public async Task<IActionResult> DownloadPaymentsSummary()
    {
        var data = await store.GetAsync();
        var workbook = PaymentSummaryWorkbook.Create(data);
        var fileName = $"payments-summary-{DateTime.Now:yyyyMMdd-HHmm}.xlsx";
        return File(
            workbook,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            fileName);
    }

    public async Task<IActionResult> DownloadClientReport(int reportYear, int? reportMonth)
    {
        var today = DateOnly.FromDateTime(DateTime.Today);
        if (reportYear < 2000 || reportYear > 2100)
        {
            reportYear = today.Year;
        }

        DateOnly startMonth;
        DateOnly endMonth;
        if (reportMonth is >= 1 and <= 12)
        {
            startMonth = new DateOnly(reportYear, reportMonth.Value, 1);
            endMonth = startMonth;
        }
        else
        {
            startMonth = new DateOnly(reportYear, 1, 1);
            var endMonthNumber = reportYear == today.Year ? today.Month : 12;
            endMonth = new DateOnly(reportYear, endMonthNumber, 1);
        }

        var data = await store.GetAsync();
        var workbook = ClientReportWorkbook.Create(data, startMonth, endMonth);
        var fileName = startMonth == endMonth
            ? $"client-report-{startMonth:yyyyMM}.xlsx"
            : $"client-report-{reportYear}.xlsx";

        return File(
            workbook,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            fileName);
    }

    public async Task<IActionResult> DownloadMonthlyPayments(string? month, string? startMonth, string? endMonth)
    {
        var currentMonth = DateOnly.FromDateTime(DateTime.Today);
        currentMonth = new DateOnly(currentMonth.Year, currentMonth.Month, 1);

        var selectedMonth = ParseMonth(month);
        var start = selectedMonth ?? ParseMonth(startMonth) ?? currentMonth;
        var end = selectedMonth ?? ParseMonth(endMonth) ?? start;
        if (end < start)
        {
            (start, end) = (end, start);
        }

        var data = await store.GetAsync();
        var workbook = start == end
            ? MonthlyPaymentsWorkbook.Create(data, start)
            : MonthlyPaymentsWorkbook.Create(data, start, end);
        var fileName = start == end
            ? $"monthly-payments-{start:yyyyMM}.xlsx"
            : $"monthly-payments-{start:yyyyMM}-to-{end:yyyyMM}.xlsx";

        return File(
            workbook,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            fileName);
    }

    public async Task<IActionResult> PaymentReceipt(int id)
    {
        var data = await store.GetAsync();
        var payment = data.Payments.FirstOrDefault(p => p.Id == id);
        if (payment is null)
        {
            return NotFound();
        }

        var receiptClient = data.Clients.FirstOrDefault(c => c.Id == payment.ClientId);
        var paymentMonth = new DateOnly(payment.PaidOn.Year, payment.PaidOn.Month, 1);
        var model = new PaymentReceiptViewModel
        {
            Payment = payment,
            Client = receiptClient,
            Settings = data.Settings,
            BillingRule = receiptClient is null ? null : BillingRules.ForClient(receiptClient, payment.PaidOn),
            CurrentBillAmount = receiptClient is null ? 0 : CurrentBillAmountForMonth(data, receiptClient, paymentMonth),
            TotalPaid = receiptClient is null ? 0 : data.Payments.Where(p => p.ClientId == receiptClient.Id).Sum(p => p.Amount)
        };

        return View(model);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> AddPayment(Payment payment, bool returnToCustomerAccount = false, bool printThermal = false)
    {
        var data = await store.GetAsync();
        payment.Id = NextId(data.Payments.Select(p => p.Id));
        if (string.IsNullOrWhiteSpace(payment.CollectedBy))
        {
            payment.CollectedBy = User.FindFirstValue("DisplayName") ?? User.Identity?.Name ?? "";
        }

        data.Payments.Add(payment);

        var client = data.Clients.FirstOrDefault(c => c.Id == payment.ClientId);
        if (client is not null)
        {
            var previousBalance = client.Balance;
            client.Balance = Math.Max(0, previousBalance - payment.Amount);
            if (payment.Amount > previousBalance)
            {
                client.Advance += payment.Amount - previousBalance;
            }
        }

        await store.SaveAsync(data);
        if (printThermal && client is not null)
        {
            return RedirectToAction(nameof(ThermalReceipt), new { id = client.Id, paymentId = payment.Id, autoPrint = true });
        }

        if (returnToCustomerAccount && client is not null)
        {
            TempData["CustomerAccountMessage"] = $"Payment recorded: PHP {payment.Amount:N0} for {client.Name}. Receipt OR-{payment.Id:000000}.";
            return RedirectToAction(nameof(CustomerAccount), new { id = client.Id, year = payment.PaidOn.Year });
        }

        return RedirectToAction(nameof(PaymentReceipt), new { id = payment.Id });
    }

    public async Task<IActionResult> Expenses()
    {
        var data = await store.GetAsync();
        return View(data.Expenses.OrderByDescending(e => e.SpentOn).ThenByDescending(e => e.Id).ToList());
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> AddExpense(Expense expense)
    {
        var data = await store.GetAsync();
        expense.Id = NextId(data.Expenses.Select(e => e.Id));
        data.Expenses.Add(expense);
        await store.SaveAsync(data);
        return RedirectToAction(nameof(Expenses));
    }

    public async Task<IActionResult> Jobs()
    {
        var data = await store.GetAsync();
        ViewBag.Clients = data.Clients.OrderBy(c => c.Name).ToList();
        ViewBag.Technicians = data.Technicians.OrderBy(t => t.Name).ToList();
        ViewBag.TechnicianOptions = BuildTechnicianAssignmentOptions(data);
        return View(data.Jobs.OrderBy(j => j.Status == "Done").ThenBy(j => j.ScheduledOn).ToList());
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> AddJob(JobTicket job)
    {
        if (string.IsNullOrWhiteSpace(job.Type))
        {
            TempData["JobError"] = "Please select a job type.";
            return RedirectToAction(nameof(Jobs));
        }

        var data = await store.GetAsync();
        var technicianOptions = BuildTechnicianAssignmentOptions(data);
        if (job.TechnicianId.HasValue && job.TechnicianId.Value > 0 && technicianOptions.All(technician => technician.Id != job.TechnicianId.Value))
        {
            TempData["JobError"] = "Please select a valid technician.";
            return RedirectToAction(nameof(Jobs));
        }

        job.Id = NextId(data.Jobs.Select(j => j.Id));
        job.Type = job.Type.Trim();
        job.TechnicianId = job.TechnicianId > 0 ? job.TechnicianId : null;
        job.Status = string.IsNullOrWhiteSpace(job.Status) ? "Open" : job.Status;
        job.Remarks = job.Remarks?.Trim() ?? "";
        data.Jobs.Add(job);
        await store.SaveAsync(data);
        return RedirectToAction(nameof(Jobs));
    }

    public async Task<IActionResult> Pppoe(string? q, string? filter, int show = 25)
    {
        var data = await store.GetAsync();
        var model = await BuildPppoeModel(data, q, filter, show, saveSync: false);
        return View(model);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SyncPppoeFromMikrotik(string? q, string? filter, int show = 25)
    {
        var data = await store.GetAsync();
        await BuildPppoeModel(data, q, filter, show, saveSync: true);
        await store.SaveAsync(data);
        return RedirectToAction(nameof(Pppoe), new { q, filter, show });
    }

    public async Task<IActionResult> Plans()
    {
        var data = await store.GetAsync();
        var model = await BuildPlansModel(data);
        return View(model);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SavePlan(ServicePlan plan)
    {
        var data = await store.GetAsync();
        plan.PlanName = plan.PlanName?.Trim() ?? "";
        plan.Type = string.IsNullOrWhiteSpace(plan.Type) ? "Prepaid" : plan.Type.Trim();

        if (plan.Id == 0)
        {
            plan.Id = NextId(data.Plans.Select(p => p.Id));
            data.Plans.Add(plan);
        }
        else
        {
            var existing = data.Plans.FirstOrDefault(p => p.Id == plan.Id);
            if (existing is null)
            {
                return NotFound();
            }

            existing.PlanName = plan.PlanName;
            existing.Price = plan.Price;
            existing.Type = plan.Type;
        }

        await store.SaveAsync(data);
        return RedirectToAction(nameof(Plans));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeletePlan(int id)
    {
        var data = await store.GetAsync();
        data.Plans.RemoveAll(p => p.Id == id);
        await store.SaveAsync(data);
        return RedirectToAction(nameof(Plans));
    }

    public async Task<IActionResult> CoverageTable()
    {
        var data = await store.GetAsync();
        ViewBag.Clients = data.Clients;
        ViewBag.KnownAreas = data.Clients
            .Select(c => c.Area)
            .Where(a => !string.IsNullOrWhiteSpace(a))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Order(StringComparer.OrdinalIgnoreCase)
            .ToList();
        return View(data.CoverageAreas.OrderBy(a => a.AreaName).ToList());
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SaveCoverage(CoverageArea area)
    {
        var data = await store.GetAsync();
        area.AreaName = area.AreaName?.Trim() ?? "";
        if (string.IsNullOrWhiteSpace(area.AreaName))
        {
            return RedirectToAction(nameof(CoverageTable));
        }

        if (area.Id == 0)
        {
            area.Id = NextId(data.CoverageAreas.Select(a => a.Id));
            data.CoverageAreas.Add(area);
        }
        else
        {
            var existing = data.CoverageAreas.FirstOrDefault(a => a.Id == area.Id);
            if (existing is null)
            {
                return NotFound();
            }

            existing.AreaName = area.AreaName;
        }

        await store.SaveAsync(data);
        return RedirectToAction(nameof(CoverageTable));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteCoverage(int id)
    {
        var data = await store.GetAsync();
        data.CoverageAreas.RemoveAll(a => a.Id == id);
        await store.SaveAsync(data);
        return RedirectToAction(nameof(CoverageTable));
    }

    public async Task<IActionResult> CoverageMap()
    {
        var data = await store.GetAsync();
        return View(data);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> AddNapLocation(NapLocation nap)
    {
        var data = await store.GetAsync();
        nap.Id = NextId(data.NapLocations.Select(n => n.Id));
        nap.Name = string.IsNullOrWhiteSpace(nap.Name) ? $"NAP {nap.Id}" : nap.Name.Trim();
        nap.AreaName = nap.AreaName?.Trim() ?? "";
        nap.Remarks = nap.Remarks?.Trim() ?? "";
        data.NapLocations.Add(nap);
        await store.SaveAsync(data);
        return RedirectToAction(nameof(CoverageMap));
    }

    public async Task<IActionResult> PonManagement()
    {
        var data = await store.GetAsync();
        return View(BuildPonManagementModel(data));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SaveOlt(OltDevice olt)
    {
        var data = await store.GetAsync();
        olt.OltName = olt.OltName?.Trim() ?? "";
        olt.Technology = string.IsNullOrWhiteSpace(olt.Technology) ? "Gpon" : olt.Technology.Trim();
        olt.Site = olt.Site?.Trim() ?? "";
        olt.ManagementUrl = NormalizeUrl(olt.ManagementUrl);
        olt.Username = olt.Username?.Trim() ?? "";
        olt.Password = olt.Password ?? "";

        if (olt.Id == 0)
        {
            olt.Id = NextId(data.OltDevices.Select(o => o.Id));
            data.OltDevices.Add(olt);
            EnsureOltPonPorts(data, olt, olt.TotalPonPorts, []);
            SetOltPonPortCapacities(data, olt);
        }
        else
        {
            var existing = data.OltDevices.FirstOrDefault(o => o.Id == olt.Id);
            if (existing is null)
            {
                return NotFound();
            }

            existing.OltName = olt.OltName;
            existing.Technology = olt.Technology;
            existing.Site = olt.Site;
            existing.ManagementUrl = olt.ManagementUrl;
            existing.Username = olt.Username;
            existing.Password = string.IsNullOrWhiteSpace(olt.Password) ? existing.Password : olt.Password;
            existing.TotalPonPorts = olt.TotalPonPorts;
            EnsureOltPonPorts(data, existing, existing.TotalPonPorts, []);
            SetOltPonPortCapacities(data, existing);
        }

        await store.SaveAsync(data);
        return RedirectToAction(nameof(PonManagement));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SaveOltPonPort(OltPonPort portRow)
    {
        var data = await store.GetAsync();
        var olt = data.OltDevices.FirstOrDefault(o => o.Id == portRow.OltDeviceId);
        if (olt is null)
        {
            return NotFound();
        }

        portRow.PonPort = NormalizePonPortLabel(portRow.PonPort);
        portRow.CustomerCapacity = CustomerCapacityForTechnology(olt.Technology);
        portRow.TotalNap = Math.Max(0, portRow.TotalNap);

        if (string.IsNullOrWhiteSpace(portRow.PonPort))
        {
            portRow.PonPort = $"PON{Math.Max(1, data.OltPonPorts.Count(p => p.OltDeviceId == portRow.OltDeviceId) + 1)}";
        }

        var existing = portRow.Id == 0
            ? data.OltPonPorts.FirstOrDefault(p => p.OltDeviceId == portRow.OltDeviceId
                && NormalizePonPortLabel(p.PonPort).Equals(portRow.PonPort, StringComparison.OrdinalIgnoreCase))
            : data.OltPonPorts.FirstOrDefault(p => p.Id == portRow.Id);

        if (existing is null)
        {
            portRow.Id = NextId(data.OltPonPorts.Select(p => p.Id));
            data.OltPonPorts.Add(portRow);
        }
        else
        {
            existing.OltDeviceId = olt.Id;
            existing.PonPort = portRow.PonPort;
            existing.CustomerCapacity = portRow.CustomerCapacity;
            existing.TotalNap = portRow.TotalNap;
        }

        await store.SaveAsync(data);
        TempData["PonSuccess"] = $"{olt.OltName} {portRow.PonPort} capacity updated.";
        return RedirectToAction(nameof(PonManagement));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SyncOltClients()
    {
        var data = await store.GetAsync();
        var syncedAt = DateTime.Now;
        var addedClients = 0;
        var failedMessages = new List<string>();

        foreach (var olt in data.OltDevices.Where(o => !string.IsNullOrWhiteSpace(o.ManagementUrl)))
        {
            var result = await oltWebClient.GetOnuClientsAsync(olt, HttpContext.RequestAborted);
            if (!result.IsSuccess)
            {
                failedMessages.Add($"{olt.OltName}: {result.ErrorMessage}");
                continue;
            }

            data.OltOnuClients.RemoveAll(c => c.OltDeviceId == olt.Id);
            foreach (var client in result.Clients)
            {
                client.Id = NextId(data.OltOnuClients.Select(c => c.Id));
                client.SyncedAt = syncedAt;
                data.OltOnuClients.Add(client);
            }

            if (result.PonPortCount > 0)
            {
                olt.TotalPonPorts = result.PonPortCount;
            }

            EnsureOltPonPorts(data, olt, result.PonPortCount, result.Clients.Select(c => c.PonPort));
            addedClients += result.Clients.Count;
        }

        await store.SaveAsync(data);

        if (failedMessages.Count == 0)
        {
            TempData["PonSuccess"] = $"OLT clients synced. {addedClients:N0} ONU records loaded.";
        }
        else
        {
            TempData["PonError"] = $"Synced {addedClients:N0} ONU records. Failed: {string.Join("; ", failedMessages)}";
        }

        return RedirectToAction(nameof(PonManagement));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteOlt(int id)
    {
        var data = await store.GetAsync();
        data.OltDevices.RemoveAll(o => o.Id == id);
        data.OltPonPorts.RemoveAll(p => p.OltDeviceId == id);
        data.OltOnuClients.RemoveAll(c => c.OltDeviceId == id);
        await store.SaveAsync(data);
        return RedirectToAction(nameof(PonManagement));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteOltPonPort(int id)
    {
        var data = await store.GetAsync();
        data.OltPonPorts.RemoveAll(p => p.Id == id);
        await store.SaveAsync(data);
        return RedirectToAction(nameof(PonManagement));
    }

    public async Task<IActionResult> AssignCollector()
    {
        var data = await store.GetAsync();
        return View(data);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SaveCollectorAssignment(CollectorAssignment assignment)
    {
        var data = await store.GetAsync();
        assignment.Id = NextId(data.CollectorAssignments.Select(a => a.Id));
        assignment.CollectorName = assignment.CollectorName?.Trim() ?? "";
        if (string.IsNullOrWhiteSpace(assignment.CollectorName) && assignment.CollectorUserId is int collectorUserId)
        {
            var collector = data.UserAccounts.FirstOrDefault(u => u.Id == collectorUserId);
            assignment.CollectorName = collector is null
                ? ""
                : string.IsNullOrWhiteSpace(collector.DisplayName) ? collector.Username : collector.DisplayName;
        }
        assignment.AreaName = assignment.AreaName?.Trim() ?? "";
        assignment.Remarks = assignment.Remarks?.Trim() ?? "";
        data.CollectorAssignments.Add(assignment);
        await store.SaveAsync(data);
        return RedirectToAction(nameof(AssignCollector));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteCollectorAssignment(int id)
    {
        var data = await store.GetAsync();
        data.CollectorAssignments.RemoveAll(a => a.Id == id);
        await store.SaveAsync(data);
        return RedirectToAction(nameof(AssignCollector));
    }

    public async Task<IActionResult> CollectionHistory()
    {
        var data = await store.GetAsync();
        ViewBag.ClientMap = data.Clients.ToDictionary(c => c.Id);
        return View(data.Payments.OrderByDescending(p => p.PaidOn).ThenByDescending(p => p.Id).ToList());
    }

    public async Task<IActionResult> Tickets()
    {
        var data = await store.GetAsync();
        ViewBag.TechnicianOptions = BuildTechnicianAssignmentOptions(data);
        return View(data);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> AddTicket(SupportTicket ticket)
    {
        if (ticket.ClientId <= 0)
        {
            TempData["TicketError"] = "Please choose a client for the ticket.";
            return RedirectToAction(nameof(Tickets));
        }

        if (string.IsNullOrWhiteSpace(ticket.Subject))
        {
            TempData["TicketError"] = "Please select a ticket reason.";
            return RedirectToAction(nameof(Tickets));
        }

        var data = await store.GetAsync();
        var technicianOptions = BuildTechnicianAssignmentOptions(data);
        var assignedTechnician = ticket.AssignedTechnicianId.HasValue
            ? technicianOptions.FirstOrDefault(technician => technician.Id == ticket.AssignedTechnicianId.Value)
            : null;

        if (assignedTechnician is null)
        {
            TempData["TicketError"] = "Please select a technician.";
            return RedirectToAction(nameof(Tickets));
        }

        ticket.Id = NextId(data.Tickets.Select(t => t.Id));
        ticket.Subject = string.IsNullOrWhiteSpace(ticket.Subject) ? "Customer ticket" : ticket.Subject.Trim();
        ticket.Type = string.IsNullOrWhiteSpace(ticket.Type) ? "Repair" : ticket.Type.Trim();
        ticket.Priority = string.IsNullOrWhiteSpace(ticket.Priority) ? "Normal" : ticket.Priority.Trim();
        ticket.Status = string.IsNullOrWhiteSpace(ticket.Status) ? "Open" : ticket.Status.Trim();
        ticket.AssignedTechnicianId = assignedTechnician.Id;
        ticket.AssignedTo = assignedTechnician.Name;
        ticket.Remarks = ticket.Remarks?.Trim() ?? "";
        data.Tickets.Add(ticket);
        await store.SaveAsync(data);
        return RedirectToAction(nameof(Tickets));
    }

    public async Task<IActionResult> JobHistory()
    {
        var data = await store.GetAsync();
        ViewBag.Clients = data.Clients.ToDictionary(c => c.Id);
        ViewBag.Technicians = data.Technicians.ToDictionary(t => t.Id);
        ViewBag.TechnicianOptions = BuildTechnicianAssignmentOptions(data).ToDictionary(t => t.Id, t => t.Name);
        return View(data.Jobs.OrderByDescending(j => j.CompletedAt ?? j.ScheduledOn.ToDateTime(TimeOnly.MinValue)).ToList());
    }

    public async Task<IActionResult> SettingsIntegrations()
    {
        var data = await store.GetAsync();
        ViewBag.Users = data.UserAccounts.OrderBy(u => u.Username).ToList();
        return View(data.Settings);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SettingsIntegrations(SystemSettings settings, IFormFile? companyLogoFile, IFormFile? browserLogoFile)
    {
        var data = await store.GetAsync();
        try
        {
            data.Settings = await MergeSettingsAsync(data.Settings, settings, companyLogoFile, browserLogoFile);
            await store.SaveAsync(data);
            TempData["SettingsSuccess"] = "Settings saved.";
        }
        catch (InvalidOperationException ex)
        {
            TempData["SettingsError"] = ex.Message;
        }

        return RedirectToAction(nameof(SettingsIntegrations));
    }

    public IActionResult BackupDatabase()
    {
        var dataPath = BillingDataPath();
        if (!System.IO.File.Exists(dataPath))
        {
            TempData["SettingsError"] = "Database file was not found.";
            return RedirectToAction(nameof(SettingsIntegrations));
        }

        var fileName = $"billing-data-backup-{DateTime.Now:yyyyMMdd-HHmmss}.json";
        return PhysicalFile(dataPath, "application/json", fileName);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> RestoreDatabase(IFormFile? backupFile)
    {
        if (backupFile is null || backupFile.Length == 0)
        {
            TempData["SettingsError"] = "Please choose a backup JSON file to restore.";
            return RedirectToAction(nameof(SettingsIntegrations));
        }

        if (!Path.GetExtension(backupFile.FileName).Equals(".json", StringComparison.OrdinalIgnoreCase))
        {
            TempData["SettingsError"] = "Restore file must be a .json database backup.";
            return RedirectToAction(nameof(SettingsIntegrations));
        }

        try
        {
            await using var stream = backupFile.OpenReadStream();
            var restored = await JsonSerializer.DeserializeAsync<BillingData>(
                stream,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

            if (restored is null)
            {
                TempData["SettingsError"] = "The selected backup file is empty or invalid.";
                return RedirectToAction(nameof(SettingsIntegrations));
            }

            BackupBillingDataFile("manual-restore");
            await store.SaveAsync(restored);
            await auditLogger.LogAsync(HttpContext, "Admin.RestoreDatabase", $"Restored database from '{backupFile.FileName}'.");
            TempData["SettingsSuccess"] = "Database restored successfully. The previous database was backed up first.";
        }
        catch (JsonException ex)
        {
            TempData["SettingsError"] = $"Restore failed: invalid JSON file. {ex.Message}";
        }
        catch (Exception ex)
        {
            TempData["SettingsError"] = $"Restore failed: {ex.Message}";
        }

        return RedirectToAction(nameof(SettingsIntegrations));
    }

    public async Task<IActionResult> Reports()
    {
        var data = await store.GetAsync();
        return View(data);
    }

    public async Task<IActionResult> Technicians()
    {
        var data = await store.GetAsync();
        return View(data.Technicians.OrderBy(t => t.Name).ToList());
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> AddTechnician(Technician technician)
    {
        var data = await store.GetAsync();
        technician.Id = NextId(data.Technicians.Select(t => t.Id));
        data.Technicians.Add(technician);
        await store.SaveAsync(data);
        return RedirectToAction(nameof(Technicians));
    }

    public async Task<IActionResult> Settings()
    {
        var data = await store.GetAsync();
        return View(data.Settings);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Settings(SystemSettings settings)
    {
        var data = await store.GetAsync();
        data.Settings.CompanyName = string.IsNullOrWhiteSpace(settings.CompanyName) ? "Billing System" : settings.CompanyName.Trim();
        data.Settings.SystemDisplayName = string.IsNullOrWhiteSpace(settings.SystemDisplayName)
            ? data.Settings.CompanyName
            : settings.SystemDisplayName.Trim();
        data.Settings.CompanyLogoUrl = string.IsNullOrWhiteSpace(settings.CompanyLogoUrl) ? data.Settings.CompanyLogoUrl : settings.CompanyLogoUrl.Trim();
        data.Settings.BrowserLogoUrl = string.IsNullOrWhiteSpace(settings.BrowserLogoUrl) ? data.Settings.BrowserLogoUrl : settings.BrowserLogoUrl.Trim();
        data.Settings.MonthlyDueDay = settings.MonthlyDueDay;
        data.Settings.Currency = string.IsNullOrWhiteSpace(settings.Currency) ? "PHP" : settings.Currency.Trim();
        data.Settings.SmsReminderTemplate = settings.SmsReminderTemplate?.Trim() ?? "";
        await store.SaveAsync(data);
        return RedirectToAction(nameof(Settings));
    }

    public async Task<IActionResult> SmsReminders()
    {
        var data = await store.GetAsync();
        return View(data.Clients.Where(c => c.Balance > 0).OrderByDescending(c => c.Balance).ToList());
    }

    public async Task<IActionResult> Users()
    {
        var data = await store.GetAsync();
        ViewBag.Roles = AccountRoles;
        ViewBag.Technicians = data.Technicians.Where(t => t.IsActive).OrderBy(t => t.Name).ToList();
        return View(data.UserAccounts.OrderBy(u => u.Username).ToList());
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> AddUser(UserAccount account)
    {
        var data = await store.GetAsync();
        var username = account.Username?.Trim() ?? "";

        if (string.IsNullOrWhiteSpace(username))
        {
            TempData["UserError"] = "Username is required.";
            return RedirectToAction(nameof(Users));
        }

        if (data.UserAccounts.Any(user => user.Username.Equals(username, StringComparison.OrdinalIgnoreCase)))
        {
            TempData["UserError"] = $"Username '{username}' already exists.";
            return RedirectToAction(nameof(Users));
        }

        if (string.IsNullOrWhiteSpace(account.Password))
        {
            TempData["UserError"] = "Password is required for a new user.";
            return RedirectToAction(nameof(Users));
        }

        account.Id = NextId(data.UserAccounts.Select(user => user.Id));
        account.Username = username;
        account.Password = account.Password.Trim();
        account.Role = NormalizeAccountRole(account.Role);
        account.DisplayName = string.IsNullOrWhiteSpace(account.DisplayName) ? username : account.DisplayName.Trim();
        account.TechnicianId = NormalizeTechnicianId(account.TechnicianId);

        data.UserAccounts.Add(account);
        await store.SaveAsync(data);
        await auditLogger.LogAsync(HttpContext, "Admin.AddUser", $"Created user '{account.Username}' with role '{account.Role}'.");

        TempData["UserSuccess"] = $"User '{account.Username}' was created.";
        return RedirectToAction(nameof(Users));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> EditUser(UserAccount account)
    {
        var data = await store.GetAsync();
        var existing = data.UserAccounts.FirstOrDefault(user => user.Id == account.Id);
        if (existing is null)
        {
            return NotFound();
        }

        var username = account.Username?.Trim() ?? "";
        if (string.IsNullOrWhiteSpace(username))
        {
            TempData["UserError"] = "Username is required.";
            return RedirectToAction(nameof(Users));
        }

        if (data.UserAccounts.Any(user => user.Id != account.Id && user.Username.Equals(username, StringComparison.OrdinalIgnoreCase)))
        {
            TempData["UserError"] = $"Username '{username}' already exists.";
            return RedirectToAction(nameof(Users));
        }

        var newRole = NormalizeAccountRole(account.Role);
        var wouldRemoveLastAdmin =
            existing.Role.Equals("Admin", StringComparison.OrdinalIgnoreCase) &&
            (!newRole.Equals("Admin", StringComparison.OrdinalIgnoreCase) || !account.IsActive) &&
            data.UserAccounts.Count(user => user.Id != existing.Id && user.IsActive && user.Role.Equals("Admin", StringComparison.OrdinalIgnoreCase)) == 0;

        if (wouldRemoveLastAdmin)
        {
            TempData["UserError"] = "Keep at least one active admin account.";
            return RedirectToAction(nameof(Users));
        }

        var previousUsername = existing.Username;
        existing.Username = username;
        existing.DisplayName = string.IsNullOrWhiteSpace(account.DisplayName) ? username : account.DisplayName.Trim();
        existing.Role = newRole;
        existing.TechnicianId = NormalizeTechnicianId(account.TechnicianId);
        existing.IsActive = account.IsActive;

        if (!string.IsNullOrWhiteSpace(account.Password))
        {
            existing.Password = account.Password.Trim();
        }

        await store.SaveAsync(data);
        await auditLogger.LogAsync(HttpContext, "Admin.EditUser", $"Edited user '{previousUsername}' as '{existing.Username}' with role '{existing.Role}'.");

        TempData["UserSuccess"] = $"User '{existing.Username}' was updated.";
        return RedirectToAction(nameof(Users));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteUser(int id)
    {
        var data = await store.GetAsync();
        var existing = data.UserAccounts.FirstOrDefault(user => user.Id == id);
        if (existing is null)
        {
            return NotFound();
        }

        var currentUserId = int.TryParse(User.FindFirstValue(ClaimTypes.NameIdentifier), out var userId) ? userId : 0;
        if (existing.Id == currentUserId)
        {
            TempData["UserError"] = "You cannot delete the account you are currently using.";
            return RedirectToAction(nameof(Users));
        }

        if (existing.Role.Equals("Admin", StringComparison.OrdinalIgnoreCase) &&
            data.UserAccounts.Count(user => user.IsActive && user.Role.Equals("Admin", StringComparison.OrdinalIgnoreCase)) <= 1)
        {
            TempData["UserError"] = "Keep at least one active admin account.";
            return RedirectToAction(nameof(Users));
        }

        data.UserAccounts.Remove(existing);
        await store.SaveAsync(data);
        await auditLogger.LogAsync(HttpContext, "Admin.DeleteUser", $"Deleted user '{existing.Username}' with role '{existing.Role}'.");

        TempData["UserSuccess"] = $"User '{existing.Username}' was deleted.";
        return RedirectToAction(nameof(Users));
    }

    public async Task<IActionResult> ActivityLogs()
    {
        var data = await store.GetAsync();
        return View(data.ActivityLogs.OrderByDescending(log => log.OccurredAt).ThenByDescending(log => log.Id).Take(500).ToList());
    }

    private static IReadOnlyList<TechnicianAssignmentOption> BuildTechnicianAssignmentOptions(BillingData data)
    {
        var options = new Dictionary<int, string>();

        foreach (var technician in data.Technicians.Where(technician => technician.IsActive))
        {
            if (technician.Id > 0)
            {
                options[technician.Id] = string.IsNullOrWhiteSpace(technician.Name)
                    ? $"Technician #{technician.Id}"
                    : technician.Name.Trim();
            }
        }

        foreach (var user in data.UserAccounts.Where(user =>
            user.IsActive &&
            user.Role.Equals("Technician", StringComparison.OrdinalIgnoreCase)))
        {
            var technicianId = EffectiveTechnicianId(user);
            if (technicianId > 0)
            {
                options[technicianId] = string.IsNullOrWhiteSpace(user.DisplayName)
                    ? user.Username.Trim()
                    : user.DisplayName.Trim();
            }
        }

        return options
            .Select(option => new TechnicianAssignmentOption(option.Key, option.Value))
            .OrderBy(option => option.Name)
            .ThenBy(option => option.Id)
            .ToList();
    }

    private static int EffectiveTechnicianId(UserAccount user)
    {
        if (user.TechnicianId is > 0)
        {
            return user.TechnicianId.Value;
        }

        return user.Role.Equals("Technician", StringComparison.OrdinalIgnoreCase) ? user.Id : 0;
    }

    private static string NormalizeAccountRole(string? role)
    {
        var match = AccountRoles.FirstOrDefault(allowedRole => allowedRole.Equals(role?.Trim(), StringComparison.OrdinalIgnoreCase));
        return match ?? "User";
    }

    private static int? NormalizeTechnicianId(int? technicianId)
    {
        return technicianId is > 0 ? technicianId : null;
    }

    private static async Task<PlansPageViewModel> BuildPlansModel(BillingData data)
    {
        var localPlans = data.Plans
            .OrderBy(p => p.Price)
            .ThenBy(p => p.PlanName)
            .ToList();
        var profiles = new List<MikrotikPlanProfileViewModel>();
        var settings = data.Settings;
        var isConnected = false;
        var connectionMessage = "";

        if (string.IsNullOrWhiteSpace(settings.MikrotikHost)
            || string.IsNullOrWhiteSpace(settings.MikrotikApiUser)
            || string.IsNullOrWhiteSpace(settings.MikrotikApiPassword))
        {
            connectionMessage = "MikroTik settings are incomplete.";
        }
        else
        {
            try
            {
                using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(12));
                var client = new MikrotikRouterOsClient(
                    settings.MikrotikHost,
                    settings.MikrotikApiPort <= 0 ? 8728 : settings.MikrotikApiPort,
                    settings.MikrotikApiUser,
                    settings.MikrotikApiPassword);
                var snapshot = await client.GetSnapshotAsync(timeout.Token);
                profiles = BuildMikrotikPlanProfiles(snapshot);
                isConnected = true;
                connectionMessage = "Connected to MikroTik";
            }
            catch (Exception ex)
            {
                connectionMessage = $"MikroTik connection failed: {ex.Message}";
            }
        }

        var matchedProfileNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var planRows = localPlans
            .Select(plan =>
            {
                var profile = profiles.FirstOrDefault(p => PlanMatchesProfile(plan, p.Name));
                if (profile is not null)
                {
                    matchedProfileNames.Add(profile.Name);
                }

                return new PlanWithCountsViewModel
                {
                    Plan = plan,
                    LocalClientCount = data.Clients.Count(c =>
                        c.PlanAmount == plan.Price
                        && (c.BillingType ?? "").Equals(plan.Type, StringComparison.OrdinalIgnoreCase)),
                    MikrotikUserCount = profile?.TotalUsers ?? 0,
                    MikrotikProfile = profile
                };
            })
            .ToList();

        return new PlansPageViewModel
        {
            Plans = planRows,
            MikrotikProfiles = profiles
                .Select(profile => CopyMikrotikProfile(profile, matchedProfileNames.Contains(profile.Name)))
                .ToList(),
            IsMikrotikConnected = isConnected,
            MikrotikConnectionMessage = connectionMessage,
            RouterHost = settings.MikrotikHost,
            TotalLocalClients = data.Clients.Count,
            TotalMikrotikUsers = profiles.Sum(profile => profile.TotalUsers)
        };
    }

    private static List<MikrotikPlanProfileViewModel> BuildMikrotikPlanProfiles(MikrotikSnapshot snapshot)
    {
        var activeUsernames = snapshot.ActiveSessions
            .Where(session => !string.IsNullOrWhiteSpace(session.Name))
            .Select(session => session.Name.Trim())
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var profilesByName = snapshot.Profiles
            .Where(profile => !string.IsNullOrWhiteSpace(profile.Name))
            .GroupBy(profile => profile.Name.Trim(), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
        var secretsByProfile = snapshot.Secrets
            .Where(secret => !string.IsNullOrWhiteSpace(secret.Profile))
            .GroupBy(secret => secret.Profile.Trim(), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.OrdinalIgnoreCase);

        return profilesByName.Keys
            .Concat(secretsByProfile.Keys)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(name => name)
            .Select(name =>
            {
                profilesByName.TryGetValue(name, out var profile);
                secretsByProfile.TryGetValue(name, out var secrets);
                secrets ??= [];

                var disabledUsers = secrets.Count(secret => secret.Disabled);
                var onlineUsers = secrets.Count(secret => !secret.Disabled && activeUsernames.Contains(secret.Name));

                return new MikrotikPlanProfileViewModel
                {
                    Name = name,
                    RateLimit = profile?.RateLimit ?? "",
                    LocalAddress = profile?.LocalAddress ?? "",
                    RemoteAddress = profile?.RemoteAddress ?? "",
                    DnsServer = profile?.DnsServer ?? "",
                    OnlyOne = profile?.OnlyOne ?? "",
                    Comment = profile?.Comment ?? "",
                    TotalUsers = secrets.Count,
                    OnlineUsers = onlineUsers,
                    DisabledUsers = disabledUsers,
                    OfflineUsers = Math.Max(0, secrets.Count - onlineUsers - disabledUsers)
                };
            })
            .ToList();
    }

    private static MikrotikPlanProfileViewModel CopyMikrotikProfile(MikrotikPlanProfileViewModel profile, bool isMatched)
    {
        return new MikrotikPlanProfileViewModel
        {
            Name = profile.Name,
            RateLimit = profile.RateLimit,
            LocalAddress = profile.LocalAddress,
            RemoteAddress = profile.RemoteAddress,
            DnsServer = profile.DnsServer,
            OnlyOne = profile.OnlyOne,
            Comment = profile.Comment,
            TotalUsers = profile.TotalUsers,
            OnlineUsers = profile.OnlineUsers,
            OfflineUsers = profile.OfflineUsers,
            DisabledUsers = profile.DisabledUsers,
            IsMatchedToLocalPlan = isMatched
        };
    }

    private static bool PlanMatchesProfile(ServicePlan plan, string profileName)
    {
        if (string.IsNullOrWhiteSpace(profileName))
        {
            return false;
        }

        var planKey = NormalizePlanKey(plan.PlanName);
        var profileKey = NormalizePlanKey(profileName);
        if (planKey.Length >= 3 && profileKey.Length >= 3
            && (planKey.Equals(profileKey, StringComparison.OrdinalIgnoreCase)
                || planKey.Contains(profileKey, StringComparison.OrdinalIgnoreCase)
                || profileKey.Contains(planKey, StringComparison.OrdinalIgnoreCase)))
        {
            return true;
        }

        return ProfileHasPriceToken(profileName, plan.Price);
    }

    private static string NormalizePlanKey(string value)
    {
        return new string((value ?? "")
            .Where(char.IsLetterOrDigit)
            .Select(char.ToLowerInvariant)
            .ToArray());
    }

    private static bool ProfileHasPriceToken(string profileName, decimal price)
    {
        if (price <= 0)
        {
            return false;
        }

        foreach (Match match in Regex.Matches(profileName, @"\d+(?:[\.,]\d+)?"))
        {
            if (decimal.TryParse(
                    match.Value.Replace(',', '.'),
                    NumberStyles.Number,
                    CultureInfo.InvariantCulture,
                    out var value)
                && Math.Abs(value - price) < 0.01m)
            {
                return true;
            }
        }

        return false;
    }

    private static string NormalizeUrl(string? url)
    {
        var value = url?.Trim() ?? "";
        if (string.IsNullOrWhiteSpace(value)
            || value.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
            || value.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        {
            return value;
        }

        return $"https://{value}";
    }

    private static async Task<PppoeManagementViewModel> BuildPppoeModel(
        BillingData data,
        string? q,
        string? filter,
        int show,
        bool saveSync)
    {
        var settings = data.Settings;
        var query = q?.Trim() ?? "";
        var selectedFilter = string.IsNullOrWhiteSpace(filter) ? "All" : filter.Trim();
        var selectedShow = show is 10 or 25 or 50 or 100 ? show : 25;

        if (string.IsNullOrWhiteSpace(settings.MikrotikHost)
            || string.IsNullOrWhiteSpace(settings.MikrotikApiUser)
            || string.IsNullOrWhiteSpace(settings.MikrotikApiPassword))
        {
            var localRows = ApplyPppoeFilters(BuildLocalPppoeRows(data), query, selectedFilter)
                .Take(selectedShow)
                .Select((row, index) => row with { Number = index + 1 })
                .ToList();

            return new PppoeManagementViewModel
            {
                IsConnected = false,
                ConnectionMessage = "MikroTik settings are incomplete.",
                RouterHost = settings.MikrotikHost,
                Query = query,
                Filter = selectedFilter,
                Show = selectedShow,
                Accounts = localRows,
                TotalUsers = data.PppoeUsers.Count,
                ActiveUsers = localRows.Count(r => r.Status == "Online"),
                OfflineUsers = localRows.Count(r => r.Status == "Offline"),
                DisabledUsers = localRows.Count(r => r.Status == "Disabled"),
                TotalUsageGb = localRows.Sum(r => r.UsageGb)
            };
        }

        try
        {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(12));
            var client = new MikrotikRouterOsClient(
                settings.MikrotikHost,
                settings.MikrotikApiPort <= 0 ? 8728 : settings.MikrotikApiPort,
                settings.MikrotikApiUser,
                settings.MikrotikApiPassword);
            var snapshot = await client.GetSnapshotAsync(timeout.Token);
            var allRows = BuildPppoeRowsFromSnapshot(data, snapshot);

            if (saveSync)
            {
                SyncPppoeRowsToStore(data, allRows);
            }

            var filteredRows = ApplyPppoeFilters(allRows, query, selectedFilter)
                .Take(selectedShow)
                .Select((row, index) => row with { Number = index + 1 })
                .ToList();

            return new PppoeManagementViewModel
            {
                IsConnected = true,
                ConnectionMessage = "Connected to MikroTik",
                RouterHost = settings.MikrotikHost,
                RouterIdentity = snapshot.Identity,
                RouterAddress = snapshot.Address,
                Version = snapshot.Version,
                BoardName = snapshot.BoardName,
                Uptime = snapshot.Uptime,
                CpuLoad = snapshot.CpuLoad,
                FreeMemory = snapshot.FreeMemory,
                TotalMemory = snapshot.TotalMemory,
                Query = query,
                Filter = selectedFilter,
                Show = selectedShow,
                Accounts = filteredRows,
                TotalUsers = allRows.Count,
                ActiveUsers = allRows.Count(r => r.Status == "Online"),
                OfflineUsers = allRows.Count(r => r.Status == "Offline"),
                DisabledUsers = allRows.Count(r => r.Status == "Disabled"),
                TotalUsageGb = allRows.Sum(r => r.UsageGb)
            };
        }
        catch (Exception ex)
        {
            var localRows = ApplyPppoeFilters(BuildLocalPppoeRows(data), query, selectedFilter)
                .Take(selectedShow)
                .Select((row, index) => row with { Number = index + 1 })
                .ToList();

            return new PppoeManagementViewModel
            {
                IsConnected = false,
                ConnectionMessage = $"MikroTik connection failed: {ex.Message}",
                RouterHost = settings.MikrotikHost,
                Query = query,
                Filter = selectedFilter,
                Show = selectedShow,
                Accounts = localRows,
                TotalUsers = localRows.Count,
                ActiveUsers = localRows.Count(r => r.Status == "Online"),
                OfflineUsers = localRows.Count(r => r.Status == "Offline"),
                DisabledUsers = localRows.Count(r => r.Status == "Disabled"),
                TotalUsageGb = localRows.Sum(r => r.UsageGb)
            };
        }
    }

    private static List<PppoeAccountViewModel> BuildPppoeRowsFromSnapshot(BillingData data, MikrotikSnapshot snapshot)
    {
        var clientsByPppoe = data.Clients
            .Where(c => !string.IsNullOrWhiteSpace(c.PppoeUsername))
            .GroupBy(c => c.PppoeUsername.Trim(), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
        var activeByName = snapshot.ActiveSessions
            .GroupBy(s => s.Name.Trim(), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        return snapshot.Secrets
            .OrderBy(s => s.Name)
            .Select(secret =>
            {
                activeByName.TryGetValue(secret.Name, out var active);
                clientsByPppoe.TryGetValue(secret.Name, out var client);

                var status = secret.Disabled ? "Disabled" : active is not null ? "Online" : "Offline";
                var usageGb = active is null ? 0 : BytesToGb(active.BytesIn + active.BytesOut);

                return new PppoeAccountViewModel
                {
                    ClientId = client?.Id,
                    CustomerName = client?.Name ?? "",
                    AccountNumber = client?.AccountNumber ?? "",
                    Username = secret.Name,
                    Address = FirstNonEmpty(active?.Address ?? "", secret.RemoteAddress),
                    CallerId = FirstNonEmpty(active?.CallerId ?? "", secret.CallerId),
                    Profile = secret.Profile,
                    LastSeen = active is not null ? $"Online {active.Uptime}" : status,
                    Status = status,
                    UsageGb = usageGb,
                    IsAssigned = client is not null
                };
            })
            .ToList();
    }

    private static List<PppoeAccountViewModel> BuildLocalPppoeRows(BillingData data)
    {
        return data.Clients
            .Where(c => !string.IsNullOrWhiteSpace(c.PppoeUsername))
            .OrderBy(c => c.PppoeUsername)
            .Select(client =>
            {
                var pppoe = data.PppoeUsers.FirstOrDefault(p => p.ClientId == client.Id);
                var status = NormalizePppoeStatus(pppoe?.Status ?? client.Status);
                return new PppoeAccountViewModel
                {
                    ClientId = client.Id,
                    CustomerName = client.Name,
                    AccountNumber = client.AccountNumber,
                    Username = pppoe?.Username ?? client.PppoeUsername,
                    Address = pppoe?.IpAddress ?? "",
                    CallerId = "",
                    Profile = $"Php {client.PlanAmount:N0}",
                    LastSeen = pppoe?.LastSeenAt?.ToString("MMM dd, yyyy h:mm tt", CultureInfo.InvariantCulture) ?? status,
                    Status = status,
                    IsAssigned = true
                };
            })
            .ToList();
    }

    private static List<PppoeAccountViewModel> ApplyPppoeFilters(
        IEnumerable<PppoeAccountViewModel> rows,
        string query,
        string filter)
    {
        var filtered = rows;
        if (!string.IsNullOrWhiteSpace(query))
        {
            filtered = filtered.Where(row =>
                row.Username.Contains(query, StringComparison.OrdinalIgnoreCase)
                || row.CustomerName.Contains(query, StringComparison.OrdinalIgnoreCase)
                || row.Profile.Contains(query, StringComparison.OrdinalIgnoreCase)
                || row.Status.Contains(query, StringComparison.OrdinalIgnoreCase)
                || row.Address.Contains(query, StringComparison.OrdinalIgnoreCase));
        }

        filtered = filter switch
        {
            "Online" => filtered.Where(row => row.Status == "Online"),
            "Offline" => filtered.Where(row => row.Status == "Offline"),
            "Disabled" => filtered.Where(row => row.Status == "Disabled"),
            "Assigned" => filtered.Where(row => row.IsAssigned),
            "Unassigned" => filtered.Where(row => !row.IsAssigned),
            _ => filtered
        };

        return filtered.ToList();
    }

    private static void SyncPppoeRowsToStore(BillingData data, IEnumerable<PppoeAccountViewModel> rows)
    {
        foreach (var row in rows.Where(r => r.ClientId.HasValue))
        {
            var clientId = row.ClientId!.Value;
            var existing = data.PppoeUsers.FirstOrDefault(p => p.ClientId == clientId);
            if (existing is null)
            {
                existing = new PppoeUser
                {
                    Id = NextId(data.PppoeUsers.Select(p => p.Id)),
                    ClientId = clientId
                };
                data.PppoeUsers.Add(existing);
            }

            existing.Username = row.Username;
            existing.Status = row.Status;
            existing.IpAddress = row.Address;
            if (row.Status == "Online")
            {
                existing.LastSeenAt = DateTime.Now;
            }

            var client = data.Clients.FirstOrDefault(c => c.Id == clientId);
            if (client is not null && row.Status == "Disabled")
            {
                client.Status = "DC";
            }
        }
    }

    private static string NormalizePppoeStatus(string value)
    {
        if (value.Equals("Active", StringComparison.OrdinalIgnoreCase)
            || value.Equals("Online", StringComparison.OrdinalIgnoreCase))
        {
            return "Online";
        }

        if (value.Equals("DC", StringComparison.OrdinalIgnoreCase)
            || value.Equals("Disabled", StringComparison.OrdinalIgnoreCase)
            || value.Contains("disconnect", StringComparison.OrdinalIgnoreCase))
        {
            return "Disabled";
        }

        return "Offline";
    }

    private static decimal BytesToGb(long bytes)
    {
        return Math.Round((decimal)bytes / 1024 / 1024 / 1024, 2, MidpointRounding.AwayFromZero);
    }

    private static string FirstNonEmpty(params string[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
    }

    private static PonManagementViewModel BuildPonManagementModel(BillingData data)
    {
        var oltDevices = data.OltDevices
            .OrderBy(o => o.Site)
            .ThenBy(o => o.OltName)
            .ToList();
        var oltById = oltDevices.ToDictionary(o => o.Id);
        var clients = data.OltOnuClients.ToList();
        var ponPorts = BuildEffectiveOltPonPorts(data, oltDevices);

        var ponRows = ponPorts
            .Where(ponPort => oltById.ContainsKey(ponPort.OltDeviceId))
            .Select(ponPort =>
            {
                var olt = oltById[ponPort.OltDeviceId];
                var portClients = clients
                    .Where(client => client.OltDeviceId == ponPort.OltDeviceId
                        && SamePonPort(client.PonPort, ponPort.PonPort))
                    .ToList();
                var usedPorts = CountDistinctOnus(portClients);
                var displayPonPort = WithTechnologyCustomerCapacity(ponPort, olt.Technology);
                var totalCapacity = displayPonPort.CustomerCapacity;

                return new OltPonMonitoringRowViewModel
                {
                    Olt = olt,
                    PonPort = displayPonPort,
                    TotalCapacity = totalCapacity,
                    UsedPorts = usedPorts,
                    AssignedCustomers = CountAssignedCustomers(portClients),
                    PonUsagePercent = UsagePercent(usedPorts, totalCapacity)
                };
            })
            .OrderBy(row => row.Olt.Site)
            .ThenBy(row => row.Olt.OltName)
            .ThenBy(row => PonSortKey(row.PonPort.PonPort))
            .ThenBy(row => row.PonPort.PonPort)
            .ToList();

        var oltRows = oltDevices
            .Select(olt =>
            {
                var rows = ponRows.Where(row => row.Olt.Id == olt.Id).ToList();
                var totalCapacity = rows.Sum(row => row.TotalCapacity);
                var usedPorts = rows.Sum(row => row.UsedPorts);

                return new OltMonitoringRowViewModel
                {
                    Olt = olt,
                    TotalCapacity = totalCapacity,
                    UsedPorts = usedPorts,
                    AssignedCustomers = rows.Sum(row => row.AssignedCustomers),
                    PonUsagePercent = UsagePercent(usedPorts, totalCapacity)
                };
            })
            .ToList();

        var overallCapacity = ponRows.Sum(row => row.TotalCapacity);
        var overallUsed = ponRows.Sum(row => row.UsedPorts);

        return new PonManagementViewModel
        {
            OltDevices = oltDevices,
            OltRows = oltRows,
            PonRows = ponRows,
            LastSyncedAt = clients.Count == 0 ? null : clients.Max(c => c.SyncedAt),
            TotalPonPorts = ponRows.Count,
            TotalCapacity = overallCapacity,
            UsedPorts = overallUsed,
            AssignedCustomers = ponRows.Sum(row => row.AssignedCustomers),
            PonUsagePercent = UsagePercent(overallUsed, overallCapacity)
        };
    }

    private static List<OltPonPort> BuildEffectiveOltPonPorts(BillingData data, IReadOnlyList<OltDevice> oltDevices)
    {
        var rows = new List<OltPonPort>();

        foreach (var olt in oltDevices)
        {
            var configuredRows = data.OltPonPorts
                .Where(ponPort => ponPort.OltDeviceId == olt.Id)
                .GroupBy(ponPort => NormalizePonPortLabel(ponPort.PonPort), StringComparer.OrdinalIgnoreCase)
                .Where(group => !string.IsNullOrWhiteSpace(group.Key))
                .Select(group =>
                {
                    var ponPort = group.First();
                    return new OltPonPort
                    {
                        Id = ponPort.Id,
                        OltDeviceId = olt.Id,
                        PonPort = group.Key,
                        CustomerCapacity = CustomerCapacityForTechnology(olt.Technology),
                        TotalNap = ponPort.TotalNap
                    };
                })
                .ToList();

            rows.AddRange(configuredRows);
            var configuredLabels = configuredRows
                .Select(row => row.PonPort)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var discoveredLabels = data.OltOnuClients
                .Where(client => client.OltDeviceId == olt.Id)
                .Select(client => NormalizePonPortLabel(client.PonPort))
                .Where(label => !string.IsNullOrWhiteSpace(label))
                .Concat(Enumerable.Range(1, Math.Min(Math.Max(olt.TotalPonPorts, 0), 128)).Select(port => $"PON{port}"))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(PonSortKey)
                .ThenBy(label => label);

            foreach (var label in discoveredLabels)
            {
                if (configuredLabels.Contains(label))
                {
                    continue;
                }

                rows.Add(new OltPonPort
                {
                    OltDeviceId = olt.Id,
                    PonPort = label,
                    CustomerCapacity = CustomerCapacityForTechnology(olt.Technology),
                    TotalNap = 0
                });
            }
        }

        return rows;
    }

    private static void EnsureOltPonPorts(BillingData data, OltDevice olt, int ponPortCount, IEnumerable<string> discoveredPorts)
    {
        var labels = discoveredPorts
            .Select(NormalizePonPortLabel)
            .Where(label => !string.IsNullOrWhiteSpace(label))
            .Concat(Enumerable.Range(1, Math.Min(Math.Max(ponPortCount, 0), 128)).Select(port => $"PON{port}"))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(PonSortKey)
            .ThenBy(label => label)
            .ToList();

        foreach (var label in labels)
        {
            var existing = data.OltPonPorts.FirstOrDefault(ponPort => ponPort.OltDeviceId == olt.Id
                && NormalizePonPortLabel(ponPort.PonPort).Equals(label, StringComparison.OrdinalIgnoreCase));
            if (existing is not null)
            {
                existing.CustomerCapacity = CustomerCapacityForTechnology(olt.Technology);
                continue;
            }

            data.OltPonPorts.Add(new OltPonPort
            {
                Id = NextId(data.OltPonPorts.Select(ponPort => ponPort.Id)),
                OltDeviceId = olt.Id,
                PonPort = label,
                CustomerCapacity = CustomerCapacityForTechnology(olt.Technology),
                TotalNap = 0
            });
        }

        if (labels.Count > olt.TotalPonPorts)
        {
            olt.TotalPonPorts = labels.Count;
        }
    }

    private static int CountDistinctOnus(IEnumerable<OltOnuClient> clients)
    {
        return clients
            .Select(client => FirstNonEmpty(client.OnuId, client.SerialNumber, client.Description))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Count();
    }

    private static int CountAssignedCustomers(IEnumerable<OltOnuClient> clients)
    {
        return clients.Count(client => !string.IsNullOrWhiteSpace(FirstNonEmpty(client.OnuId, client.SerialNumber, client.Description)));
    }

    private static OltPonPort WithTechnologyCustomerCapacity(OltPonPort ponPort, string technology)
    {
        return new OltPonPort
        {
            Id = ponPort.Id,
            OltDeviceId = ponPort.OltDeviceId,
            PonPort = ponPort.PonPort,
            CustomerCapacity = CustomerCapacityForTechnology(technology),
            TotalNap = Math.Max(0, ponPort.TotalNap)
        };
    }

    private static void SetOltPonPortCapacities(BillingData data, OltDevice olt)
    {
        var customerCapacity = CustomerCapacityForTechnology(olt.Technology);
        foreach (var ponPort in data.OltPonPorts.Where(ponPort => ponPort.OltDeviceId == olt.Id))
        {
            ponPort.CustomerCapacity = customerCapacity;
        }
    }

    private static int CustomerCapacityForTechnology(string? technology)
    {
        return (technology ?? "").Contains("epon", StringComparison.OrdinalIgnoreCase) ? 32 : 128;
    }

    private static decimal UsagePercent(int usedPorts, int totalCapacity)
    {
        return totalCapacity <= 0
            ? 0
            : Math.Round((decimal)usedPorts / totalCapacity * 100, 1, MidpointRounding.AwayFromZero);
    }

    private static bool SamePonPort(string left, string right)
    {
        return NormalizePonPortLabel(left).Equals(NormalizePonPortLabel(right), StringComparison.OrdinalIgnoreCase);
    }

    private static int PonSortKey(string value)
    {
        var match = Regex.Match(value ?? "", @"\d+");
        return match.Success && int.TryParse(match.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var port)
            ? port
            : int.MaxValue;
    }

    private static string NormalizePonPortLabel(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "";
        }

        var trimmed = value.Trim();
        var gponMatch = Regex.Match(trimmed, @"GPON\d+/(?<port>\d+):", RegexOptions.IgnoreCase);
        if (gponMatch.Success)
        {
            return $"PON{gponMatch.Groups["port"].Value}";
        }

        var ponMatch = Regex.Match(trimmed, @"PON\s*(?<port>\d+)", RegexOptions.IgnoreCase);
        if (ponMatch.Success)
        {
            return $"PON{ponMatch.Groups["port"].Value}";
        }

        return int.TryParse(trimmed, NumberStyles.Integer, CultureInfo.InvariantCulture, out var port)
            ? $"PON{port}"
            : trimmed.ToUpperInvariant().Replace(" ", "");
    }

    private static int NextId(IEnumerable<int> ids) => ids.DefaultIfEmpty().Max() + 1;

    private static string NextAccountNumber(IEnumerable<Client> clients)
    {
        var lastNumber = clients
            .Select(client => int.TryParse(client.AccountNumber, NumberStyles.Integer, CultureInfo.InvariantCulture, out var number) ? number : 0)
            .DefaultIfEmpty()
            .Max();

        return (lastNumber + 1).ToString(CultureInfo.InvariantCulture);
    }

    private static int AccountSortKey(string accountNumber)
    {
        return int.TryParse(accountNumber, NumberStyles.Integer, CultureInfo.InvariantCulture, out var number)
            ? number
            : int.MaxValue;
    }

    private static string DeleteCaptchaQuestion(Client client) => $"{DeleteCaptchaNumberA(client)} + {DeleteCaptchaNumberB(client)}";

    private static string DeleteCaptchaAnswer(Client client)
    {
        return (DeleteCaptchaNumberA(client) + DeleteCaptchaNumberB(client)).ToString(CultureInfo.InvariantCulture);
    }

    private static int DeleteCaptchaNumberA(Client client) => client.Id % 9 + 1;

    private static int DeleteCaptchaNumberB(Client client)
    {
        var digitSum = (client.AccountNumber ?? "").Where(char.IsDigit).Sum(ch => ch - '0');
        return digitSum % 9 + 1;
    }

    private static CustomerStatementViewModel BuildCustomerStatementModel(
        BillingData data,
        Client client,
        string statementTitle,
        int? year,
        bool includeFullYear)
    {
        var today = DateOnly.FromDateTime(DateTime.Today);
        var currentMonth = new DateOnly(today.Year, today.Month, 1);
        var payments = data.Payments
            .Where(payment => payment.ClientId == client.Id)
            .OrderByDescending(payment => payment.PaidOn)
            .ThenByDescending(payment => payment.Id)
            .ToList();
        var planChanges = data.PlanChanges
            .Where(change => change.ClientId == client.Id)
            .OrderByDescending(change => change.EffectiveMonth)
            .ThenByDescending(change => change.Id)
            .ToList();
        var monthlyBillOverrides = data.MonthlyBillOverrides
            .Where(overrideBill => overrideBill.ClientId == client.Id)
            .OrderByDescending(overrideBill => overrideBill.BillingMonth)
            .ThenByDescending(overrideBill => overrideBill.Id)
            .ToList();
        var billingMonths = BuildCustomerBillingMonths(client, payments, planChanges, monthlyBillOverrides);
        var currentBillingMonth = billingMonths.FirstOrDefault(month => month.Month == currentMonth)
            ?? billingMonths.OrderByDescending(month => month.Month).FirstOrDefault()
            ?? new CustomerBillingMonth
            {
                Month = currentMonth,
                MonthLabel = currentMonth.ToString("MMMM yyyy", CultureInfo.InvariantCulture),
                DueDate = BillingRules.ForClient(client, today).NextDueDate,
                BillAmount = client.Bills,
                Balance = client.Balance
            };
        var selectedYear = includeFullYear
            ? year is int requestedYear ? requestedYear : today.Year
            : today.Year;
        var summaryMonths = includeFullYear
            ? billingMonths.Where(month => month.Month.Year == selectedYear).OrderBy(month => month.Month).ToList()
            : billingMonths.Where(month => month.Month == currentBillingMonth.Month).OrderBy(month => month.Month).ToList();

        if (summaryMonths.Count == 0)
        {
            summaryMonths.Add(currentBillingMonth);
        }

        var monthPayments = payments
            .Where(payment => payment.PaidOn.Year == currentBillingMonth.Month.Year && payment.PaidOn.Month == currentBillingMonth.Month.Month)
            .ToList();
        var planAmount = MonthlyPlanAmount(client, currentBillingMonth.Month, monthPayments, planChanges);
        var newCharges = StatementNewCharges(currentBillingMonth);
        if (newCharges <= 0 && currentBillingMonth.BillAmount > 0)
        {
            newCharges = currentBillingMonth.BillAmount;
        }

        var totalBalanceDue = Math.Max(
            0,
            currentBillingMonth.Balance - currentBillingMonth.Advance + newCharges - currentBillingMonth.AmountPaid);
        var accountKey = StatementAccountKey(client);
        var invoiceNumber = $"INV-{accountKey}-{currentBillingMonth.Month:yyyyMM}";

        return new CustomerStatementViewModel
        {
            Client = client,
            Settings = data.Settings,
            CompanyName = StatementCompanyName,
            CompanyAddress = StatementCompanyAddress,
            CompanyContact = StatementCompanyContact,
            StatementDate = today,
            StatementNumber = $"ST-{accountKey}-{today:yyyyMMdd}",
            InvoiceNumber = invoiceNumber,
            StatementTitle = statementTitle,
            PlanLabel = $"{client.BillingType} - PHP {planAmount:N0}",
            PlanAmount = planAmount,
            PreviousBalance = currentBillingMonth.Balance,
            Credits = currentBillingMonth.AmountPaid,
            NewCharges = newCharges,
            TotalBalanceDue = totalBalanceDue,
            OpenSupportTickets = data.Tickets.Count(ticket => ticket.ClientId == client.Id && IsOpenTicket(ticket)),
            PaymentDueDate = currentBillingMonth.DueDate,
            BillSummary = summaryMonths
                .Select(month => new CustomerStatementPeriodRow
                {
                    Period = month.MonthLabel,
                    Charges = month.BillAmount,
                    Payments = month.AmountPaid,
                    Net = Math.Max(0, month.BillAmount - month.AmountPaid)
                })
                .ToList(),
            ChargeBreakdown = includeFullYear
                ? []
                : [new CustomerStatementChargeRow
                {
                    Date = currentBillingMonth.Month,
                    SalesInvoice = invoiceNumber,
                    Description = $"{client.BillingType} internet service - {currentBillingMonth.MonthLabel}",
                    Charge = newCharges
                }]
        };
    }

    private static ThermalReceiptViewModel BuildThermalReceiptModel(BillingData data, Client client, int? paymentId = null)
    {
        var today = DateOnly.FromDateTime(DateTime.Today);
        var payments = data.Payments
            .Where(payment => payment.ClientId == client.Id)
            .OrderByDescending(payment => payment.PaidOn)
            .ThenByDescending(payment => payment.Id)
            .ToList();
        var latestPayment = paymentId.HasValue
            ? payments.FirstOrDefault(payment => payment.Id == paymentId.Value)
            : null;
        latestPayment ??= payments.FirstOrDefault();
        var receiptDate = latestPayment?.PaidOn ?? today;
        var paymentMonth = new DateOnly(receiptDate.Year, receiptDate.Month, 1);
        var monthPayments = payments
            .Where(payment => payment.PaidOn.Year == receiptDate.Year && payment.PaidOn.Month == receiptDate.Month)
            .OrderBy(payment => payment.PaidOn)
            .ThenBy(payment => payment.Id)
            .ToList();
        var planChanges = data.PlanChanges.Where(change => change.ClientId == client.Id).ToList();
        var monthlyBillOverrides = data.MonthlyBillOverrides.Where(overrideBill => overrideBill.ClientId == client.Id).ToList();
        var billingMonth = BuildCustomerBillingMonths(client, payments, planChanges, monthlyBillOverrides)
            .FirstOrDefault(month => month.Month == paymentMonth);
        var planAmount = MonthlyPlanAmount(client, paymentMonth, monthPayments, planChanges);
        var previousBalance = billingMonth?.Balance ?? client.Balance;
        var advance = billingMonth?.Advance ?? client.Advance;
        var paymentAmount = latestPayment?.Amount ?? 0;
        var totalCharge = billingMonth?.BillAmount ?? Math.Max(0, previousBalance - advance + planAmount);
        var totalPaid = monthPayments.Sum(payment => payment.Amount);

        return new ThermalReceiptViewModel
        {
            Client = client,
            Payment = latestPayment,
            Settings = data.Settings,
            CompanyName = StatementCompanyName,
            CompanyAddress = StatementCompanyAddress,
            CompanyContact = ThermalSupportContact,
            ReceiptDate = today,
            DateOfPayment = receiptDate,
            DueDate = billingMonth?.DueDate ?? DueDateForBillingMonth(client, paymentMonth),
            OfficialReceiptNumber = latestPayment is null ? "-" : $"OR-{latestPayment.Id:000000}",
            AccountLabel = string.IsNullOrWhiteSpace(client.AccountNumber) ? $"Client #{client.Id}" : client.AccountNumber,
            PlanLabel = client.BillingType,
            PlanAmount = planAmount,
            PreviousBalance = previousBalance,
            Advance = advance,
            PaymentAmount = paymentAmount,
            PreviousPayment = paymentAmount,
            BalanceAfterPayment = Math.Max(0, previousBalance - paymentAmount),
            TotalCharge = totalCharge,
            TotalPaid = totalPaid,
            CurrentBalance = Math.Max(0, totalCharge - totalPaid),
            CurrentAdvance = Math.Max(0, totalPaid - totalCharge)
        };
    }

    private static decimal StatementNewCharges(CustomerBillingMonth billingMonth)
    {
        return Math.Max(0, billingMonth.BillAmount - billingMonth.Balance + billingMonth.Advance);
    }

    private static bool IsOpenTicket(SupportTicket ticket)
    {
        var status = ticket.Status ?? "";
        return !status.Equals("Closed", StringComparison.OrdinalIgnoreCase)
            && !status.Equals("Done", StringComparison.OrdinalIgnoreCase)
            && !status.Equals("Resolved", StringComparison.OrdinalIgnoreCase)
            && !status.Equals("Cancelled", StringComparison.OrdinalIgnoreCase)
            && !status.Equals("Canceled", StringComparison.OrdinalIgnoreCase);
    }

    private static string StatementAccountKey(Client client)
    {
        var value = string.IsNullOrWhiteSpace(client.AccountNumber)
            ? client.Id.ToString(CultureInfo.InvariantCulture)
            : client.AccountNumber;

        return new string(value.Where(char.IsLetterOrDigit).ToArray()).ToUpperInvariant();
    }

    private static IReadOnlyList<CustomerBillingMonth> BuildCustomerBillingMonths(
        Client client,
        IReadOnlyList<Payment> payments,
        IReadOnlyList<ClientPlanChange> planChanges,
        IReadOnlyList<ClientMonthlyBillOverride> monthlyBillOverrides)
    {
        var months = CustomerBillingHistoryMonths(client, payments, monthlyBillOverrides);

        var rows = new List<CustomerBillingMonth>();
        var carriedBalance = 0m;
        var carriedAdvance = 0m;
        foreach (var month in months)
        {
            var monthPayments = payments
                .Where(p => p.PaidOn.Year == month.Year && p.PaidOn.Month == month.Month)
                .OrderBy(p => p.PaidOn)
                .ThenBy(p => p.Id)
                .ToList();
            var amountPaid = monthPayments.Sum(p => p.Amount);
            var billOverride = MonthlyBillOverrideFor(client, month, monthlyBillOverrides);
            var openingBalance = billOverride?.Balance ?? carriedBalance;
            var openingAdvance = billOverride?.Advance ?? carriedAdvance;
            var planAmount = MonthlyPlanAmount(client, month, monthPayments, planChanges);
            var billAmount = billOverride?.BillAmount ?? Math.Max(0, openingBalance - openingAdvance + planAmount);
            var endingBalance = billOverride is null
                ? Math.Max(0, openingBalance + planAmount - openingAdvance - amountPaid)
                : Math.Max(0, billAmount - amountPaid);
            var endingAdvance = billOverride is null
                ? Math.Max(0, openingAdvance + amountPaid - openingBalance - planAmount)
                : Math.Max(0, amountPaid - billAmount);
            var status = billAmount <= 0 ? "No bill" :
                endingBalance <= 0 ? "Paid" :
                amountPaid > 0 ? "Partial" : "Unpaid";

            rows.Add(new CustomerBillingMonth
            {
                Month = month,
                MonthLabel = month.ToString("MMMM yyyy", CultureInfo.InvariantCulture),
                DueDate = DueDateForBillingMonth(client, month),
                BillAmount = billAmount,
                AmountPaid = amountPaid,
                Balance = openingBalance,
                Advance = openingAdvance,
                DiscountAmount = billOverride?.DiscountAmount ?? 0,
                DiscountNote = billOverride?.DiscountRemarks ?? "",
                Status = status,
                PaymentBreakdown = monthPayments
                    .Select(p => new CustomerBillingPaymentBreakdown
                    {
                        PaidOn = p.PaidOn,
                        Amount = p.Amount
                    })
                    .ToList(),
                PaymentDates = monthPayments.Count == 0 ? "-" : string.Join(", ", monthPayments.Select(p => p.PaidOn.ToString("MMM dd, yyyy", CultureInfo.InvariantCulture))),
                Methods = JoinDistinct(monthPayments.Select(p => NormalizePaymentMethod(p.Method))),
                References = JoinDistinct(monthPayments.Select(p => p.ReferenceNumber)),
                Collectors = JoinDistinct(monthPayments.Select(p => p.CollectedBy)),
                Remarks = JoinDistinct(monthPayments.Select(p => p.Remarks).Concat(new[] { billOverride?.DiscountRemarks ?? "", billOverride?.Remarks ?? "" }))
            });

            carriedBalance = endingBalance;
            carriedAdvance = endingAdvance;
        }

        return rows.OrderByDescending(r => r.DueDate).ToList();
    }

    private static decimal MonthlyPlanAmount(
        Client client,
        DateOnly month,
        IReadOnlyList<Payment> monthPayments,
        IEnumerable<ClientPlanChange> planChanges)
    {
        var amount = PlanAmountForBillingMonth(client, month, planChanges);
        if (amount <= 0)
        {
            amount = client.Bills;
        }

        return amount;
    }

    private static IReadOnlyDictionary<int, decimal> CurrentBillAmountsForMonth(BillingData data, DateOnly month)
    {
        return data.Clients.ToDictionary(
            client => client.Id,
            client => WholeNumberPart(PaymentPageBillAmountForMonth(data, client, month)));
    }

    private static IReadOnlyList<ClientCurrentBillRow> BuildClientCurrentBillRows(BillingData data, DateOnly month)
    {
        return data.Clients
            .OrderBy(client => AccountSortKey(client.AccountNumber))
            .ThenBy(client => client.AccountNumber ?? "")
            .Select(client =>
            {
                var previousBalance = WholeNumberPart(PreviousBalanceForMonth(data, client, month));
                var previousAdvance = WholeNumberPart(PreviousAdvanceForMonth(data, client, month));
                var referralDiscount = WholeNumberPart(ReferralDiscountForMonth(data, client, month));
                var referralNames = ReferralNamesForMonth(data, client, month);
                var currentBill = WholeNumberPart(PaymentPageBillAmountForMonth(data, client, month));
                var paidThisMonth = data.Payments
                    .Where(payment => payment.ClientId == client.Id
                        && payment.PaidOn.Year == month.Year
                        && payment.PaidOn.Month == month.Month)
                    .Sum(payment => payment.Amount);
                var balance = Math.Max(0, currentBill - paidThisMonth);
                var status = currentBill <= 0 ? "No bill" :
                    paidThisMonth >= currentBill ? "Paid" :
                    paidThisMonth > 0 ? "Partial" : "Unpaid";

                return new ClientCurrentBillRow
                {
                    Client = client,
                    DueDate = DueDateForBillingMonth(client, month),
                    PreviousBalance = previousBalance,
                    ReferralDiscount = referralDiscount,
                    ReferralNames = referralNames,
                    CurrentBill = currentBill,
                    PaidThisMonth = paidThisMonth,
                    Balance = balance,
                    Advance = previousAdvance,
                    Status = status
                };
            })
            .ToList();
    }

    private static decimal PaymentPageBillAmountForMonth(BillingData data, Client client, DateOnly month)
    {
        var billOverride = MonthlyBillOverrideFor(
            client,
            month,
            data.MonthlyBillOverrides.Where(overrideBill => overrideBill.ClientId == client.Id));
        return billOverride?.BillAmount ?? client.Bills;
    }

    private static decimal PreviousAdvanceForMonth(BillingData data, Client client, DateOnly month)
    {
        var billOverride = MonthlyBillOverrideFor(
            client,
            month,
            data.MonthlyBillOverrides.Where(overrideBill => overrideBill.ClientId == client.Id));
        if (billOverride?.Advance is decimal advance)
        {
            return advance;
        }

        var today = DateOnly.FromDateTime(DateTime.Today);
        var currentMonth = new DateOnly(today.Year, today.Month, 1);
        if (month != currentMonth)
        {
            return 0;
        }

        var installedMonth = client.DateInstalled is DateOnly installed
            ? new DateOnly(installed.Year, installed.Month, 1)
            : (DateOnly?)null;
        return installedMonth is null || installedMonth < month ? client.Advance : 0;
    }

    private static decimal ReferralDiscountForMonth(BillingData data, Client client, DateOnly month)
    {
        var billOverride = MonthlyBillOverrideFor(
            client,
            month,
            data.MonthlyBillOverrides.Where(overrideBill => overrideBill.ClientId == client.Id));
        return billOverride?.DiscountAmount ?? 0;
    }

    private static string ReferralNamesForMonth(BillingData data, Client client, DateOnly month)
    {
        var referrals = data.Referrals
            .Where(referral => referral.ReferrerClientId == client.Id)
            .ToList();
        if (referrals.Count == 0)
        {
            return "";
        }

        var billOverride = MonthlyBillOverrideFor(
            client,
            month,
            data.MonthlyBillOverrides.Where(overrideBill => overrideBill.ClientId == client.Id));
        var discountRemarks = billOverride?.DiscountRemarks ?? "";
        var monthLabel = month.ToString("MMM yyyy", CultureInfo.InvariantCulture);

        return JoinDistinct(referrals
            .Where(referral =>
                (!string.IsNullOrWhiteSpace(referral.NewClientName)
                    && discountRemarks.Contains(referral.NewClientName, StringComparison.OrdinalIgnoreCase))
                || referral.Remarks.Contains(monthLabel, StringComparison.OrdinalIgnoreCase)
                || referral.DiscountStartMonth == month)
            .Select(referral => referral.NewClientName));
    }

    private static decimal PreviousBalanceForMonth(BillingData data, Client client, DateOnly month)
    {
        var billOverride = MonthlyBillOverrideFor(
            client,
            month,
            data.MonthlyBillOverrides.Where(overrideBill => overrideBill.ClientId == client.Id));
        if (billOverride?.Balance is decimal balance)
        {
            return balance;
        }

        var today = DateOnly.FromDateTime(DateTime.Today);
        var currentMonth = new DateOnly(today.Year, today.Month, 1);
        if (month != currentMonth)
        {
            return 0;
        }

        var installedMonth = client.DateInstalled is DateOnly installed
            ? new DateOnly(installed.Year, installed.Month, 1)
            : (DateOnly?)null;
        return installedMonth is null || installedMonth < month ? client.Balance : 0;
    }

    private static decimal CurrentBillAmountForMonth(BillingData data, Client client, DateOnly month)
    {
        var billOverride = MonthlyBillOverrideFor(
            client,
            month,
            data.MonthlyBillOverrides.Where(overrideBill => overrideBill.ClientId == client.Id));
        if (billOverride is not null)
        {
            return billOverride.BillAmount;
        }

        var monthPayments = data.Payments
            .Where(payment => payment.ClientId == client.Id
                && payment.PaidOn.Year == month.Year
                && payment.PaidOn.Month == month.Month)
            .ToList();
        return MonthlyPlanAmount(
            client,
            month,
            monthPayments,
            data.PlanChanges.Where(change => change.ClientId == client.Id));
    }

    private static decimal WholeNumberPart(decimal amount)
    {
        return decimal.Truncate(amount);
    }

    private static ClientMonthlyBillOverride? MonthlyBillOverrideFor(
        Client client,
        DateOnly month,
        IEnumerable<ClientMonthlyBillOverride> monthlyBillOverrides)
    {
        return monthlyBillOverrides
            .Where(overrideBill => overrideBill.ClientId == client.Id && overrideBill.BillingMonth == month)
            .OrderByDescending(overrideBill => overrideBill.Id)
            .FirstOrDefault();
    }

    private static decimal PlanAmountForBillingMonth(Client client, DateOnly month, IEnumerable<ClientPlanChange> planChanges)
    {
        var planChange = planChanges
            .Where(change => change.ClientId == client.Id && change.EffectiveMonth <= month)
            .OrderByDescending(change => change.EffectiveMonth)
            .ThenByDescending(change => change.Id)
            .FirstOrDefault();

        return planChange?.PlanAmount ?? client.PlanAmount;
    }

    private static void UpsertMonthlyBillOverride(
        BillingData data,
        int clientId,
        DateOnly billingMonth,
        decimal billAmount,
        string remarks,
        decimal? discountAmount = null,
        string? discountRemarks = null)
    {
        var existing = data.MonthlyBillOverrides
            .FirstOrDefault(overrideBill => overrideBill.ClientId == clientId && overrideBill.BillingMonth == billingMonth);

        if (existing is not null)
        {
            existing.BillAmount = billAmount;
            existing.RecordedAt = DateTime.Now;
            existing.Remarks = remarks;
            if (discountAmount.HasValue)
            {
                existing.DiscountAmount = discountAmount.Value;
                existing.DiscountRemarks = discountRemarks ?? "";
            }

            return;
        }

        data.MonthlyBillOverrides.Add(new ClientMonthlyBillOverride
        {
            Id = NextId(data.MonthlyBillOverrides.Select(overrideBill => overrideBill.Id)),
            ClientId = clientId,
            BillingMonth = billingMonth,
            BillAmount = billAmount,
            DiscountAmount = discountAmount ?? 0,
            DiscountRemarks = discountRemarks ?? "",
            RecordedAt = DateTime.Now,
            Remarks = remarks
        });
    }

    private static DateOnly CustomerBillingStartMonth(Client client, IReadOnlyList<Payment> payments, DateOnly currentMonth)
    {
        var startMonth = client.DateInstalled is DateOnly installed
            ? new DateOnly(installed.Year, installed.Month, 1)
            : payments.Count > 0
                ? new DateOnly(payments.Min(p => p.PaidOn).Year, payments.Min(p => p.PaidOn).Month, 1)
                : currentMonth.AddMonths(-11);

        return startMonth > currentMonth ? currentMonth : startMonth;
    }

    private static IReadOnlyList<DateOnly> CustomerBillingHistoryMonths(
        Client client,
        IReadOnlyList<Payment> payments,
        IReadOnlyList<ClientMonthlyBillOverride> monthlyBillOverrides)
    {
        return payments
            .Where(payment => payment.ClientId == client.Id)
            .Select(payment => new DateOnly(payment.PaidOn.Year, payment.PaidOn.Month, 1))
            .Concat(monthlyBillOverrides
                .Where(overrideBill => overrideBill.ClientId == client.Id)
                .Select(overrideBill => overrideBill.BillingMonth))
            .Distinct()
            .Order()
            .ToList();
    }

    private static void EnsurePlanBaseline(BillingData data, Client client, IReadOnlyList<Payment> payments)
    {
        if (data.PlanChanges.Any(change => change.ClientId == client.Id))
        {
            return;
        }

        var today = DateOnly.FromDateTime(DateTime.Today);
        var currentMonth = new DateOnly(today.Year, today.Month, 1);
        var baselineMonth = CustomerBillingStartMonth(client, payments, currentMonth);
        data.PlanChanges.Add(new ClientPlanChange
        {
            Id = NextId(data.PlanChanges.Select(change => change.Id)),
            ClientId = client.Id,
            EffectiveMonth = baselineMonth,
            PlanAmount = client.PlanAmount,
            RecordedAt = DateTime.Now,
            Remarks = "Baseline plan before plan change."
        });
    }

    private static void UpsertPlanChange(BillingData data, int clientId, DateOnly effectiveMonth, decimal planAmount, string remarks)
    {
        var existing = data.PlanChanges
            .FirstOrDefault(change => change.ClientId == clientId && change.EffectiveMonth == effectiveMonth);

        if (existing is not null)
        {
            existing.PlanAmount = planAmount;
            existing.RecordedAt = DateTime.Now;
            existing.Remarks = remarks;
            return;
        }

        data.PlanChanges.Add(new ClientPlanChange
        {
            Id = NextId(data.PlanChanges.Select(change => change.Id)),
            ClientId = clientId,
            EffectiveMonth = effectiveMonth,
            PlanAmount = planAmount,
            RecordedAt = DateTime.Now,
            Remarks = remarks
        });
    }

    private static DateOnly DueDateForBillingMonth(Client client, DateOnly month)
    {
        return BillingRules.DueDateForBillingMonth(client, month);
    }

    private static string JoinDistinct(IEnumerable<string> values)
    {
        var distinct = values
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => value.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        return distinct.Count == 0 ? "-" : string.Join(", ", distinct);
    }

    private static DateOnly? ParseMonth(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        return DateOnly.TryParseExact(
            $"{value}-01",
            "yyyy-MM-dd",
            CultureInfo.InvariantCulture,
            DateTimeStyles.None,
            out var month)
            ? month
            : null;
    }

    private static IReadOnlyList<DashboardMethodSummary> BuildMethodSummaries(IReadOnlyList<Payment> payments, decimal totalCollected)
    {
        var groups = payments
            .GroupBy(p => NormalizePaymentMethod(p.Method))
            .ToDictionary(g => g.Key, g => new { Amount = g.Sum(p => p.Amount), Count = g.Count() });

        DashboardMethodSummary summary(string method, string barClass, string badgeClass)
        {
            groups.TryGetValue(method, out var group);
            var amount = group?.Amount ?? 0;
            return new DashboardMethodSummary
            {
                Method = method,
                Amount = amount,
                Count = group?.Count ?? 0,
                Percentage = totalCollected == 0 ? 0 : amount / totalCollected * 100,
                BarClass = barClass,
                BadgeClass = badgeClass
            };
        }

        return
        [
            summary("Cash", "bg-green", "bg-green-lt"),
            summary("GCash", "bg-blue", "bg-blue-lt"),
            summary("Other", "bg-secondary", "bg-secondary-lt")
        ];
    }

    private static string NormalizePaymentMethod(string method)
    {
        if (method.Contains("gcash", StringComparison.OrdinalIgnoreCase))
        {
            return "GCash";
        }

        if (method.Contains("cash", StringComparison.OrdinalIgnoreCase))
        {
            return "Cash";
        }

        return "Other";
    }
}

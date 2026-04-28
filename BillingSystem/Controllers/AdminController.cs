using System.Globalization;
using BillingSystem.Models;
using BillingSystem.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace BillingSystem.Controllers;

[Authorize(Roles = "Admin")]
public sealed class AdminController(IBillingStore store, IWebHostEnvironment environment) : Controller
{
    private const string ClearClientsConfirmation = "CLEAR ALL";

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
            "account-desc" => clients.OrderByDescending(c => AccountSortKey(c.AccountNumber)).ThenByDescending(c => c.AccountNumber ?? "").ThenBy(c => c.Name ?? ""),
            "name" => clients.OrderBy(c => c.Name ?? ""),
            "type" => clients.OrderBy(c => c.BillingType ?? "").ThenBy(c => c.Name ?? ""),
            "status" => clients.OrderBy(c => c.Status ?? "").ThenBy(c => c.Name ?? ""),
            "area" => clients.OrderBy(c => c.Area ?? "").ThenBy(c => c.Zone ?? "").ThenBy(c => c.Name ?? ""),
            "pppoe" => clients.OrderBy(c => string.IsNullOrWhiteSpace(c.PppoeUsername)).ThenBy(c => c.PppoeUsername ?? "").ThenBy(c => c.Name ?? ""),
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
        var dataPath = Path.Combine(environment.ContentRootPath, "Data", "billing-data.json");
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

    private static void ClearClientOperationalRecords(BillingData data)
    {
        data.Referrals.Clear();
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
    public async Task<IActionResult> UpdateMonthlyBill(int id, string? billingMonth, decimal billAmount, int? returnYear)
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

        UpsertMonthlyBillOverride(data, client.Id, month.Value, billAmount, "Current bill changed from customer account.");

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
            client.BillingType = string.IsNullOrWhiteSpace(client.BillingType) ? "Prepaid" : client.BillingType;
            client.Referral = NormalizeReferralText(client.Referral);
            client.Bills = ProratedBill(client.PlanAmount, client.DateInstalled.Value);
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
                $"Prorated first bill from installation date {client.DateInstalled.Value:MMM dd, yyyy}.");

            var referralResult = ApplyReferralDiscount(data, client);
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
            existing.BillingType = client.BillingType;
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

    public async Task<IActionResult> Payments()
    {
        var data = await store.GetAsync();
        ViewBag.Clients = data.Clients.OrderBy(c => c.Name).ToList();
        ViewBag.ClientMap = data.Clients.ToDictionary(c => c.Id);
        ViewBag.ClientBillingRules = data.Clients.ToDictionary(c => c.Id, c => BillingRules.ForClient(c));
        return View(data.Payments.OrderByDescending(p => p.PaidOn).ThenByDescending(p => p.Id).ToList());
    }

    public async Task<IActionResult> PaymentHistory()
    {
        var data = await store.GetAsync();
        ViewBag.ClientMap = data.Clients.ToDictionary(c => c.Id);
        ViewBag.PaymentTableTitle = "Payment history";
        return View(data.Payments.OrderByDescending(p => p.PaidOn).ThenByDescending(p => p.Id).ToList());
    }

    public async Task<IActionResult> RecentPaidBills()
    {
        var data = await store.GetAsync();
        ViewBag.ClientMap = data.Clients.ToDictionary(c => c.Id);
        ViewBag.PaymentTableTitle = "Recent Paid Bills";
        return View(data.Payments.OrderByDescending(p => p.PaidOn).ThenByDescending(p => p.Id).Take(50).ToList());
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
        var model = new PaymentReceiptViewModel
        {
            Payment = payment,
            Client = receiptClient,
            Settings = data.Settings,
            BillingRule = receiptClient is null ? null : BillingRules.ForClient(receiptClient, payment.PaidOn),
            TotalPaid = receiptClient is null ? 0 : data.Payments.Where(p => p.ClientId == receiptClient.Id).Sum(p => p.Amount)
        };

        return View(model);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> AddPayment(Payment payment)
    {
        var data = await store.GetAsync();
        payment.Id = NextId(data.Payments.Select(p => p.Id));
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
        ViewBag.Clients = data.Clients;
        return View(data.Plans.OrderBy(p => p.Price).ThenBy(p => p.PlanName).ToList());
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
        return View(data.OltDevices.OrderBy(o => o.Site).ThenBy(o => o.OltName).ToList());
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SaveOlt(OltDevice olt)
    {
        var data = await store.GetAsync();
        olt.OltName = olt.OltName?.Trim() ?? "";
        olt.Technology = string.IsNullOrWhiteSpace(olt.Technology) ? "Gpon" : olt.Technology.Trim();
        olt.Site = olt.Site?.Trim() ?? "";

        if (olt.Id == 0)
        {
            olt.Id = NextId(data.OltDevices.Select(o => o.Id));
            data.OltDevices.Add(olt);
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
            existing.TotalPonPorts = olt.TotalPonPorts;
        }

        await store.SaveAsync(data);
        return RedirectToAction(nameof(PonManagement));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteOlt(int id)
    {
        var data = await store.GetAsync();
        data.OltDevices.RemoveAll(o => o.Id == id);
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

        if (string.IsNullOrWhiteSpace(ticket.AssignedTo))
        {
            TempData["TicketError"] = "Please select a technician.";
            return RedirectToAction(nameof(Tickets));
        }

        var data = await store.GetAsync();
        ticket.Id = NextId(data.Tickets.Select(t => t.Id));
        ticket.Subject = string.IsNullOrWhiteSpace(ticket.Subject) ? "Customer ticket" : ticket.Subject.Trim();
        ticket.Type = string.IsNullOrWhiteSpace(ticket.Type) ? "Repair" : ticket.Type.Trim();
        ticket.Priority = string.IsNullOrWhiteSpace(ticket.Priority) ? "Normal" : ticket.Priority.Trim();
        ticket.Status = string.IsNullOrWhiteSpace(ticket.Status) ? "Open" : ticket.Status.Trim();
        ticket.AssignedTo = ticket.AssignedTo?.Trim() ?? "";
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
    public async Task<IActionResult> SettingsIntegrations(SystemSettings settings)
    {
        var data = await store.GetAsync();
        data.Settings = settings;
        await store.SaveAsync(data);
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
        data.Settings = settings;
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
        return View(data.UserAccounts.OrderBy(u => u.Username).ToList());
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

    private static int NextId(IEnumerable<int> ids) => ids.DefaultIfEmpty().Max() + 1;

    private static string NextAccountNumber(IEnumerable<Client> clients)
    {
        var lastNumber = clients
            .Select(client => int.TryParse(client.AccountNumber, NumberStyles.Integer, CultureInfo.InvariantCulture, out var number) ? number : 0)
            .DefaultIfEmpty()
            .Max();

        return (lastNumber + 1).ToString(CultureInfo.InvariantCulture);
    }

    private static decimal ProratedBill(decimal planAmount, DateOnly installedOn)
    {
        if (planAmount <= 0)
        {
            return 0;
        }

        var daysInMonth = DateTime.DaysInMonth(installedOn.Year, installedOn.Month);
        var billableDays = Math.Max(0, daysInMonth - installedOn.Day + 1);
        return Math.Round(planAmount / daysInMonth * billableDays, 2, MidpointRounding.AwayFromZero);
    }

    private static string ApplyReferralDiscount(BillingData data, Client newClient)
    {
        var referralText = NormalizeReferralText(newClient.Referral);
        if (referralText.Equals("INQUIRE", StringComparison.OrdinalIgnoreCase))
        {
            newClient.Referral = "INQUIRE";
            return "";
        }

        var referrer = FindReferralClient(data.Clients, newClient);
        if (referrer is null)
        {
            newClient.Referral = referralText;
            return $"{newClient.Name} was added, but referral \"{referralText}\" did not match one existing client.";
        }

        newClient.Referral = ReferralOptionText(referrer);
        var installedOn = newClient.DateInstalled ?? DateOnly.FromDateTime(DateTime.Today);
        var discountStartMonth = new DateOnly(installedOn.Year, installedOn.Month, 1).AddMonths(1);
        var discountCredit = Math.Round(referrer.PlanAmount / 2, 2, MidpointRounding.AwayFromZero);
        if (discountCredit <= 0)
        {
            return $"{newClient.Name} was added with referral to {referrer.Name}, but no discount was applied because the referrer has no plan amount.";
        }

        var remainingCredit = discountCredit;
        var appliedNotes = new List<string>();
        var discountMonth = discountStartMonth;
        var referrerPlanChanges = data.PlanChanges.Where(change => change.ClientId == referrer.Id).ToList();

        for (var guard = 0; guard < 120 && remainingCredit > 0; guard++)
        {
            var plannedBill = PlanAmountForBillingMonth(referrer, discountMonth, referrerPlanChanges);
            if (plannedBill <= 0)
            {
                plannedBill = referrer.PlanAmount > 0 ? referrer.PlanAmount : referrer.Bills;
            }

            if (plannedBill <= 0)
            {
                break;
            }

            var existingBill = MonthlyBillOverrideFor(referrer, discountMonth, data.MonthlyBillOverrides);
            var currentBill = existingBill?.BillAmount ?? plannedBill;
            if (currentBill <= 0)
            {
                discountMonth = discountMonth.AddMonths(1);
                continue;
            }

            var appliedDiscount = Math.Min(remainingCredit, currentBill);
            var note = $"Referral discount PHP {appliedDiscount:N0} for referring {newClient.Name}.";
            ApplyMonthlyReferralDiscount(data, referrer.Id, discountMonth, currentBill - appliedDiscount, appliedDiscount, note);
            appliedNotes.Add($"{discountMonth:MMM yyyy}: PHP {appliedDiscount:N0}");
            remainingCredit -= appliedDiscount;
            discountMonth = discountMonth.AddMonths(1);
        }

        var appliedAmount = discountCredit - remainingCredit;
        data.Referrals.Add(new ClientReferral
        {
            Id = NextId(data.Referrals.Select(referral => referral.Id)),
            ReferrerClientId = referrer.Id,
            NewClientId = newClient.Id,
            ReferrerName = referrer.Name,
            NewClientName = newClient.Name,
            ReferralText = newClient.Referral,
            DiscountStartMonth = discountStartMonth,
            DiscountAmount = discountCredit,
            AppliedAmount = appliedAmount,
            RecordedAt = DateTime.Now,
            Remarks = appliedNotes.Count == 0 ? "No discount was applied." : string.Join("; ", appliedNotes)
        });

        return appliedAmount > 0
            ? $"{newClient.Name} was added. {referrer.Name} gets PHP {appliedAmount:N0} referral discount starting {discountStartMonth:MMMM yyyy}."
            : $"{newClient.Name} was added with referral to {referrer.Name}, but no bill was available for discount yet.";
    }

    private static void ApplyMonthlyReferralDiscount(
        BillingData data,
        int clientId,
        DateOnly billingMonth,
        decimal billAmount,
        decimal discountAmount,
        string note)
    {
        var existing = data.MonthlyBillOverrides
            .Where(overrideBill => overrideBill.ClientId == clientId && overrideBill.BillingMonth == billingMonth)
            .OrderByDescending(overrideBill => overrideBill.Id)
            .FirstOrDefault();

        if (existing is not null)
        {
            existing.BillAmount = Math.Max(0, billAmount);
            existing.DiscountAmount += discountAmount;
            existing.DiscountRemarks = AppendNote(existing.DiscountRemarks, note);
            existing.RecordedAt = DateTime.Now;
            return;
        }

        data.MonthlyBillOverrides.Add(new ClientMonthlyBillOverride
        {
            Id = NextId(data.MonthlyBillOverrides.Select(overrideBill => overrideBill.Id)),
            ClientId = clientId,
            BillingMonth = billingMonth,
            BillAmount = Math.Max(0, billAmount),
            DiscountAmount = discountAmount,
            DiscountRemarks = note,
            RecordedAt = DateTime.Now,
            Remarks = ""
        });
    }

    private static Client? FindReferralClient(IEnumerable<Client> clients, Client newClient)
    {
        var referralText = NormalizeReferralText(newClient.Referral);
        if (referralText.Equals("INQUIRE", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var candidates = clients.Where(client => client.Id != newClient.Id).ToList();
        var optionMatch = candidates.FirstOrDefault(client =>
            ReferralOptionText(client).Equals(referralText, StringComparison.OrdinalIgnoreCase));
        if (optionMatch is not null)
        {
            return optionMatch;
        }

        var accountText = referralText.Split('-', 2)[0].Trim();
        var accountMatch = candidates.FirstOrDefault(client =>
            (client.AccountNumber ?? "").Equals(accountText, StringComparison.OrdinalIgnoreCase));
        if (accountMatch is not null)
        {
            return accountMatch;
        }

        var exactNameMatches = candidates
            .Where(client => (client.Name ?? "").Equals(referralText, StringComparison.OrdinalIgnoreCase))
            .Take(2)
            .ToList();
        if (exactNameMatches.Count == 1)
        {
            return exactNameMatches[0];
        }

        var partialNameMatches = candidates
            .Where(client => (client.Name ?? "").Contains(referralText, StringComparison.OrdinalIgnoreCase))
            .Take(2)
            .ToList();
        return partialNameMatches.Count == 1 ? partialNameMatches[0] : null;
    }

    private static string NormalizeReferralText(string? referral)
    {
        return string.IsNullOrWhiteSpace(referral) ? "INQUIRE" : referral.Trim();
    }

    private static string ReferralOptionText(Client client)
    {
        return $"{client.AccountNumber} - {client.Name}".Trim();
    }

    private static string AppendNote(string? existing, string note)
    {
        if (string.IsNullOrWhiteSpace(existing))
        {
            return note;
        }

        return existing.Contains(note, StringComparison.OrdinalIgnoreCase)
            ? existing
            : $"{existing} {note}";
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

        if (!client.BillingType.Equals("Xentronet", StringComparison.OrdinalIgnoreCase))
        {
            return amount;
        }

        var dueDate = DueDateForBillingMonth(client, month);
        return monthPayments.Any(p => p.PaidOn < dueDate)
            ? Math.Max(0, amount - BillingRules.XentronetEarlyDiscount)
            : amount;
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

    private static void UpsertMonthlyBillOverride(BillingData data, int clientId, DateOnly billingMonth, decimal billAmount, string remarks)
    {
        var existing = data.MonthlyBillOverrides
            .FirstOrDefault(overrideBill => overrideBill.ClientId == clientId && overrideBill.BillingMonth == billingMonth);

        if (existing is not null)
        {
            existing.BillAmount = billAmount;
            existing.RecordedAt = DateTime.Now;
            existing.Remarks = remarks;
            return;
        }

        data.MonthlyBillOverrides.Add(new ClientMonthlyBillOverride
        {
            Id = NextId(data.MonthlyBillOverrides.Select(overrideBill => overrideBill.Id)),
            ClientId = clientId,
            BillingMonth = billingMonth,
            BillAmount = billAmount,
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
        if (client.DateInstalled is DateOnly installed
            && installed.Year == month.Year
            && installed.Month == month.Month)
        {
            return installed;
        }

        if (client.BillingType.Equals("Postpaid", StringComparison.OrdinalIgnoreCase))
        {
            return new DateOnly(month.Year, month.Month, DateTime.DaysInMonth(month.Year, month.Month));
        }

        return new DateOnly(month.Year, month.Month, 1);
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

using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using BillingSystem.Models;

namespace BillingSystem.Services;

public sealed record ClientListImportResult(
    int TotalRows,
    int Imported,
    int SkippedDuplicate,
    int SkippedInvalid,
    int PaymentsImported,
    int MonthlyBillsImported,
    int ReferralsImported,
    int UnmatchedPayments,
    string SavedFileName);

internal sealed record ParsedClientList(
    List<Client> Clients,
    int TotalRows,
    int SkippedDuplicate,
    int SkippedInvalid);

internal sealed record ImportedPaymentHistory(
    List<Payment> Payments,
    List<ClientMonthlyBillOverride> MonthlyBillOverrides,
    int UnmatchedPayments);

internal sealed record ClientIndex(
    Dictionary<string, int> ByName,
    Dictionary<string, int> ByPppoe);

public static partial class ClientListExcelImporter
{
    private static readonly XNamespace SpreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace RelationshipNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace PackageRelationshipNs = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly IReadOnlyDictionary<string, int> MonthNumbers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
    {
        ["JAN"] = 1,
        ["JANUARY"] = 1,
        ["FEB"] = 2,
        ["FEBRUARY"] = 2,
        ["MAR"] = 3,
        ["MARCH"] = 3,
        ["APR"] = 4,
        ["APRIL"] = 4,
        ["MAY"] = 5,
        ["JUN"] = 6,
        ["JUNE"] = 6,
        ["JUL"] = 7,
        ["JULY"] = 7,
        ["AUG"] = 8,
        ["AUGUST"] = 8,
        ["SEP"] = 9,
        ["SEPTEMBER"] = 9,
        ["OCT"] = 10,
        ["OCTOBER"] = 10,
        ["NOV"] = 11,
        ["NOVEMBER"] = 11,
        ["DEC"] = 12,
        ["DECEMBER"] = 12
    };

    public static ClientListImportResult ReplaceClients(
        BillingData data,
        Stream workbookStream,
        int importYear,
        string savedFileName)
    {
        var workbook = ReadWorkbook(workbookStream);
        var parsed = ParseClientList(workbook, importYear, savedFileName, []);
        data.Clients = parsed.Clients;
        var history = ImportPaymentHistory(workbook, data.Clients);
        var activeBillingMonth = LatestBillingMonth(history.MonthlyBillOverrides, importYear);
        ApplyImportedProratedFirstBills(data.Clients, history.MonthlyBillOverrides, activeBillingMonth);
        data.Payments = history.Payments;
        data.MonthlyBillOverrides = history.MonthlyBillOverrides;
        data.PlanChanges = [];
        data.Referrals.Clear();
        var referralsImported = ReferralBillingService.ApplyReferralDiscounts(data, parsed.Clients);

        return new ClientListImportResult(
            parsed.TotalRows,
            parsed.Clients.Count,
            parsed.SkippedDuplicate,
            parsed.SkippedInvalid,
            history.Payments.Count,
            history.MonthlyBillOverrides.Count,
            referralsImported,
            history.UnmatchedPayments,
            savedFileName);
    }

    public static ClientListImportResult ImportNewClients(
        BillingData data,
        Stream workbookStream,
        int importYear,
        string savedFileName)
    {
        var workbook = ReadWorkbook(workbookStream);
        var parsed = ParseClientList(workbook, importYear, savedFileName, ExistingClientKeys(data.Clients));
        var nextId = data.Clients.Select(client => client.Id).DefaultIfEmpty().Max() + 1;
        foreach (var client in parsed.Clients)
        {
            client.Id = nextId++;
            if (string.IsNullOrWhiteSpace(client.AccountNumber))
            {
                client.AccountNumber = client.Id.ToString(CultureInfo.InvariantCulture);
            }

            data.Clients.Add(client);
        }
        ApplyImportedProratedFirstBills(parsed.Clients, data.MonthlyBillOverrides, FallbackBillingMonth(importYear));
        var referralsImported = ReferralBillingService.ApplyReferralDiscounts(data, parsed.Clients);

        return new ClientListImportResult(
            parsed.TotalRows,
            parsed.Clients.Count,
            parsed.SkippedDuplicate,
            parsed.SkippedInvalid,
            0,
            0,
            referralsImported,
            0,
            savedFileName);
    }

    private static ParsedClientList ParseClientList(
        IReadOnlyDictionary<string, List<List<object?>>> workbook,
        int importYear,
        string savedFileName,
        HashSet<string> existingKeys)
    {
        var rows = workbook.FirstOrDefault(sheet => sheet.Key.Equals("CLIENTS LIST", StringComparison.OrdinalIgnoreCase)).Value;
        if (rows is null)
        {
            throw new InvalidOperationException("CLIENTS LIST sheet was not found in the uploaded workbook.");
        }

        var headerIndex = rows.FindIndex(row => HasHeader(row, "Name") && HasHeader(row, "Plan"));
        if (headerIndex < 0)
        {
            throw new InvalidOperationException("CLIENTS LIST sheet must include at least Name and Plan columns.");
        }

        var headers = HeaderIndexes(rows[headerIndex]);
        var clients = new List<Client>();
        var nextId = 1;
        var skippedDuplicate = 0;
        var skippedInvalid = 0;
        var totalRows = 0;

        foreach (var row in rows.Skip(headerIndex + 1))
        {
            var name = Text(Get(row, headers, "Name"));
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            totalRows++;
            var accountNumber = Text(Get(row, headers, "Number"));
            var pppoe = Text(Get(row, headers, "PPPoE"));
            var key = ClientKey(accountNumber, pppoe, name);
            if (string.IsNullOrWhiteSpace(key))
            {
                skippedInvalid++;
                continue;
            }

            if (existingKeys.Contains(key))
            {
                skippedDuplicate++;
                continue;
            }

            var assignedId = nextId++;
            var area = Text(Get(row, headers, "Area"));
            var zone = Text(Get(row, headers, "Zone"));
            var facebook = Text(Get(row, headers, "FACEBOOK"));
            var status = ImportedClientStatus(Text(GetAny(row, headers, "Status", "Active", "Mode")));
            var billingType = BillingTypeValue(GetAny(row, headers, "Type", "Billing Type", "BillingType", "Account Type"));
            var dateInstalled = DateValue(GetAny(row, headers, "Date Installed", "Date", "Installed", "Installation Date"));
            var planAmount = Money(Get(row, headers, "Plan"));
            var referral = ReferralBillingService.NormalizeReferralText(Text(GetAny(
                row,
                headers,
                "Referral",
                "Refferal",
                "Referal",
                "Referred By",
                "ReferredBy",
                "Referral By",
                "Referral Name",
                "Client Referral")));

            clients.Add(new Client
            {
                Id = assignedId,
                AccountNumber = string.IsNullOrWhiteSpace(accountNumber) ? assignedId.ToString(CultureInfo.InvariantCulture) : accountNumber,
                DateInstalled = dateInstalled,
                Status = status,
                BillingType = billingType,
                PlanAmount = planAmount,
                Area = area,
                Zone = zone,
                Name = name,
                PppoeUsername = pppoe,
                Contact = facebook,
                FacebookAccount = facebook,
                Referral = referral,
                Balance = Money(Get(row, headers, "Balance")),
                Advance = Money(Get(row, headers, "Advance")),
                Bills = Money(Get(row, headers, "Bills")),
                Address = string.Join(" ", new[] { area, zone }.Where(value => !string.IsNullOrWhiteSpace(value))),
                Remarks = $"Imported from {savedFileName} for {importYear}."
            });
            existingKeys.Add(key);
        }

        return new ParsedClientList(clients, totalRows, skippedDuplicate, skippedInvalid);
    }

    private static ImportedPaymentHistory ImportPaymentHistory(
        IReadOnlyDictionary<string, List<List<object?>>> workbook,
        List<Client> clients)
    {
        var clientIndex = BuildClientIndex(clients);
        var monthlyBillOverrides = ImportMonthlyBills(workbook, clients, clientIndex);
        var payments = ImportCashPayments(workbook, clientIndex);
        payments.AddRange(ImportGcashPayments(workbook, clientIndex));

        payments = payments
            .OrderBy(payment => payment.PaidOn)
            .ThenBy(payment => payment.Id)
            .ToList();

        for (var index = 0; index < payments.Count; index++)
        {
            payments[index].Id = index + 1;
        }

        return new ImportedPaymentHistory(
            payments,
            monthlyBillOverrides,
            payments.Count(payment => payment.ClientId == 0));
    }

    private static void ApplyImportedProratedFirstBills(
        IReadOnlyList<Client> clients,
        List<ClientMonthlyBillOverride> monthlyBillOverrides,
        DateOnly activeBillingMonth)
    {
        var nextOverrideId = monthlyBillOverrides.Select(overrideBill => overrideBill.Id).DefaultIfEmpty().Max() + 1;
        foreach (var client in clients)
        {
            client.BillingType = BillingRules.NormalizeBillingType(client.BillingType);
            if (client.DateInstalled is not DateOnly installed || client.PlanAmount <= 0)
            {
                continue;
            }

            var firstMonth = MonthStart(installed);
            var firstBill = BillingRules.ProratedFirstBill(client.PlanAmount, installed, client.BillingType);
            var dueDate = BillingRules.FirstBillDueDate(installed, client.BillingType);
            var existingOverride = monthlyBillOverrides
                .Where(overrideBill => overrideBill.ClientId == client.Id && overrideBill.BillingMonth == firstMonth)
                .OrderByDescending(overrideBill => overrideBill.Id)
                .FirstOrDefault();

            if (existingOverride is not null)
            {
                if (firstMonth == activeBillingMonth)
                {
                    existingOverride.BillAmount = firstBill;
                }

                existingOverride.RecordedAt = DateTime.Now;
                existingOverride.Remarks = ProratedImportRemarks(existingOverride.Remarks, installed, dueDate);
            }
            else if (firstMonth == activeBillingMonth)
            {
                monthlyBillOverrides.Add(new ClientMonthlyBillOverride
                {
                    Id = nextOverrideId++,
                    ClientId = client.Id,
                    BillingMonth = firstMonth,
                    BillAmount = firstBill,
                    RecordedAt = DateTime.Now,
                    Remarks = ProratedImportRemarks("", installed, dueDate)
                });
            }

            if (firstMonth == activeBillingMonth)
            {
                client.Bills = firstBill;
                if (client.Balance <= 0)
                {
                    client.Balance = Math.Max(0, firstBill - client.Advance);
                }
            }
        }
    }

    private static DateOnly LatestBillingMonth(IReadOnlyList<ClientMonthlyBillOverride> monthlyBillOverrides, int importYear)
    {
        return monthlyBillOverrides
            .Select(overrideBill => overrideBill.BillingMonth)
            .DefaultIfEmpty(FallbackBillingMonth(importYear))
            .Max();
    }

    private static DateOnly FallbackBillingMonth(int importYear)
    {
        return new DateOnly(importYear, DateTime.Today.Month, 1);
    }

    private static string ProratedImportRemarks(string remarks, DateOnly installed, DateOnly dueDate)
    {
        var note = $"Prorated first bill from installation date {installed:MMM dd, yyyy}; due {dueDate:MMM dd, yyyy}.";
        if (string.IsNullOrWhiteSpace(remarks))
        {
            return note;
        }

        return remarks.Contains("Prorated first bill", StringComparison.OrdinalIgnoreCase)
            ? remarks
            : $"{remarks} {note}";
    }

    private static List<ClientMonthlyBillOverride> ImportMonthlyBills(
        IReadOnlyDictionary<string, List<List<object?>>> workbook,
        List<Client> clients,
        ClientIndex clientIndex)
    {
        var clientById = clients.ToDictionary(client => client.Id);
        var overrides = new List<ClientMonthlyBillOverride>();
        DateOnly? latestMonth = null;
        var latestRows = new List<(int ClientId, Dictionary<string, List<int>> Headers, IReadOnlyList<object?> Row)>();

        foreach (var (sheetName, rows) in workbook)
        {
            if (!BillsSheetRegex().IsMatch(sheetName))
            {
                continue;
            }

            var month = SheetMonth(sheetName);
            if (month is null)
            {
                continue;
            }

            var headerIndex = -1;
            Dictionary<string, List<int>>? headers = null;
            for (var index = 0; index < Math.Min(rows.Count, 10); index++)
            {
                var possibleHeaders = HeaderIndexesMulti(rows[index]);
                if (possibleHeaders.ContainsKey("NAME") && possibleHeaders.ContainsKey("BILLS"))
                {
                    headerIndex = index;
                    headers = possibleHeaders;
                    break;
                }
            }

            if (headerIndex < 0 || headers is null)
            {
                continue;
            }

            var nameColumn = HeaderColumn(headers, "NAME");
            var pppoeColumn = HeaderColumn(headers, "PPPOE");
            var statusColumn = HeaderColumn(headers, "STATUS");
            var planColumn = HeaderColumn(headers, "PLAN");
            var billColumn = HeaderColumn(headers, "BILLS");
            var balanceColumn = HeaderColumn(headers, "BALANCE");
            var advanceColumn = HeaderColumn(headers, "ADVANCE");
            var monthRows = new List<(int ClientId, Dictionary<string, List<int>> Headers, IReadOnlyList<object?> Row)>();

            foreach (var row in rows.Skip(headerIndex + 1))
            {
                var clientName = Text(Cell(row, nameColumn));
                if (string.IsNullOrWhiteSpace(clientName))
                {
                    continue;
                }

                var clientId = MatchClient(clientName, Text(Cell(row, pppoeColumn)), clientIndex);
                if (clientId == 0)
                {
                    continue;
                }

                clientById.TryGetValue(clientId, out var matchedClient);
                var plan = Money(Cell(row, planColumn));
                var balance = Money(Cell(row, balanceColumn));
                var advance = Money(Cell(row, advanceColumn));
                var bill = Math.Max(0, balance + plan - advance);
                var effectivePlan = plan > 0 ? plan : matchedClient?.PlanAmount ?? 0;
                DateOnly? firstBillInstalled = null;
                if (matchedClient?.DateInstalled is DateOnly installed
                    && MonthStart(installed) == month.Value
                    && effectivePlan > 0)
                {
                    firstBillInstalled = installed;
                    bill = BillingRules.ProratedFirstBill(effectivePlan, installed, matchedClient!.BillingType);
                }
                else if (bill <= 0 && billColumn is not null)
                {
                    bill = Money(Cell(row, billColumn));
                }

                var status = Text(Cell(row, statusColumn));
                if (bill > 0 || IsDisconnectedStatus(status))
                {
                    overrides.Add(new ClientMonthlyBillOverride
                    {
                        Id = overrides.Count + 1,
                        ClientId = clientId,
                        BillingMonth = month.Value,
                        BillAmount = bill,
                        Balance = balance,
                        Advance = advance,
                        RecordedAt = DateTime.Now,
                        Remarks = firstBillInstalled is DateOnly proratedInstalled
                            ? ProratedImportRemarks(
                                $"Imported from {sheetName}.",
                                proratedInstalled,
                                BillingRules.FirstBillDueDate(proratedInstalled, matchedClient!.BillingType))
                            : $"Imported from {sheetName}."
                    });
                }

                if (plan > 0 && clientById.TryGetValue(clientId, out var planClient))
                {
                    planClient.PlanAmount = plan;
                }

                monthRows.Add((clientId, headers, row));
            }

            if (latestMonth is null || month.Value > latestMonth.Value)
            {
                latestMonth = month.Value;
                latestRows = monthRows;
            }
        }

        foreach (var (clientId, headers, row) in latestRows)
        {
            if (!clientById.TryGetValue(clientId, out var client))
            {
                continue;
            }

            var status = Text(Cell(row, HeaderColumn(headers, "STATUS")));
            if (!string.IsNullOrWhiteSpace(status))
            {
                client.Status = ImportedClientStatus(status);
            }

            var billColumn = HeaderColumn(headers, "BILLS");
            var balanceColumn = HeaderColumn(headers, "BALANCE", 1) ?? HeaderColumn(headers, "BALANCE");
            var advanceColumn = HeaderColumn(headers, "ADVANCE", 1) ?? HeaderColumn(headers, "ADVANCE");

            if (billColumn is not null)
            {
                client.Bills = Money(Cell(row, billColumn));
            }

            if (balanceColumn is not null)
            {
                client.Balance = Money(Cell(row, balanceColumn));
            }

            if (advanceColumn is not null)
            {
                client.Advance = Money(Cell(row, advanceColumn));
            }
        }

        return overrides;
    }

    private static List<Payment> ImportCashPayments(
        IReadOnlyDictionary<string, List<List<object?>>> workbook,
        ClientIndex clientIndex)
    {
        var payments = new List<Payment>();
        foreach (var (sheetName, rows) in workbook)
        {
            if (!CashSheetRegex().IsMatch(sheetName))
            {
                continue;
            }

            var fallbackDate = SheetMonth(sheetName) ?? new DateOnly(DateTime.Today.Year, DateTime.Today.Month, 1);
            var lastDate = fallbackDate;
            foreach (var row in rows.Skip(5))
            {
                var paidOn = DateValue(Cell(row, 3));
                if (paidOn is not null)
                {
                    lastDate = paidOn.Value;
                }

                var clientName = Text(Cell(row, 6));
                var amount = PositiveMoney(Cell(row, 7));
                if (string.IsNullOrWhiteSpace(clientName) || amount is null)
                {
                    continue;
                }

                var clientId = MatchClient(clientName, "", clientIndex);
                var remarks = $"Imported from {sheetName}";
                if (clientId == 0)
                {
                    remarks += $"; unmatched client: {clientName}";
                }

                payments.Add(new Payment
                {
                    Id = payments.Count + 1,
                    ClientId = clientId,
                    PaidOn = paidOn ?? lastDate,
                    Amount = amount.Value,
                    Method = "Cash",
                    ReferenceNumber = "",
                    CollectedBy = "",
                    Remarks = remarks
                });
            }
        }

        return payments;
    }

    private static List<Payment> ImportGcashPayments(
        IReadOnlyDictionary<string, List<List<object?>>> workbook,
        ClientIndex clientIndex)
    {
        var payments = new List<Payment>();
        foreach (var (sheetName, rows) in workbook)
        {
            if (!GcashSheetRegex().IsMatch(sheetName))
            {
                continue;
            }

            var fallbackDate = SheetMonth(sheetName) ?? new DateOnly(DateTime.Today.Year, DateTime.Today.Month, 1);
            var lastDate = fallbackDate;
            foreach (var row in rows.Skip(6))
            {
                var paidOn = DateValue(Cell(row, 6));
                if (paidOn is not null)
                {
                    lastDate = paidOn.Value;
                }

                var clientName = Text(Cell(row, 9));
                var amount = PositiveMoney(Cell(row, 11));
                if (string.IsNullOrWhiteSpace(clientName) || amount is null)
                {
                    continue;
                }

                var clientId = MatchClient(clientName, "", clientIndex);
                var remarks = $"Imported from {sheetName}";
                if (clientId == 0)
                {
                    remarks += $"; unmatched client: {clientName}";
                }

                payments.Add(new Payment
                {
                    Id = payments.Count + 1,
                    ClientId = clientId,
                    PaidOn = paidOn ?? lastDate,
                    Amount = amount.Value,
                    Method = "GCash",
                    ReferenceNumber = Text(Cell(row, 8)),
                    CollectedBy = Text(Cell(row, 7)),
                    Remarks = remarks
                });
            }
        }

        return payments;
    }

    private static ClientIndex BuildClientIndex(IEnumerable<Client> clients)
    {
        var byName = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var byPppoe = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var client in clients)
        {
            var normalizedName = NormalizeName(client.Name);
            if (!string.IsNullOrWhiteSpace(normalizedName))
            {
                byName.TryAdd(normalizedName, client.Id);
            }

            if (!string.IsNullOrWhiteSpace(client.PppoeUsername))
            {
                byPppoe.TryAdd(client.PppoeUsername.Trim(), client.Id);
            }
        }

        return new ClientIndex(byName, byPppoe);
    }

    private static int MatchClient(string name, string pppoe, ClientIndex clientIndex)
    {
        if (!string.IsNullOrWhiteSpace(pppoe) && clientIndex.ByPppoe.TryGetValue(pppoe.Trim(), out var pppoeClientId))
        {
            return pppoeClientId;
        }

        var normalizedName = NormalizeName(name);
        if (clientIndex.ByName.TryGetValue(normalizedName, out var clientId))
        {
            return clientId;
        }

        var trimmed = PaymentMonthSuffixRegex().Replace(normalizedName, "").Trim();
        trimmed = SpaceRegex().Replace(trimmed, " ");
        return clientIndex.ByName.TryGetValue(trimmed, out var trimmedClientId) ? trimmedClientId : 0;
    }

    private static bool IsDisconnectedStatus(string status)
    {
        return status.Equals("DC", StringComparison.OrdinalIgnoreCase)
            || status.Equals("CUT", StringComparison.OrdinalIgnoreCase)
            || status.Equals("DISCONNECTED", StringComparison.OrdinalIgnoreCase);
    }

    private static string ImportedClientStatus(string status)
    {
        var value = status.Trim();
        if (string.IsNullOrWhiteSpace(value))
        {
            return "Active";
        }

        if (value.Equals("Active", StringComparison.OrdinalIgnoreCase))
        {
            return "Active";
        }

        return IsDisconnectedStatus(value) || value.Contains("disconnect", StringComparison.OrdinalIgnoreCase)
            ? "DC"
            : value;
    }

    private static string BillingTypeValue(object? value)
    {
        var text = Text(value);
        if (string.IsNullOrWhiteSpace(text))
        {
            return "Prepaid";
        }

        if (text.Contains("post", StringComparison.OrdinalIgnoreCase))
        {
            return "Postpaid";
        }

        if (text.Contains("xentro", StringComparison.OrdinalIgnoreCase))
        {
            return "Xentronet";
        }

        if (text.Contains("pre", StringComparison.OrdinalIgnoreCase))
        {
            return "Prepaid";
        }

        return BillingRules.NormalizeBillingType(text);
    }

    private static HashSet<string> ExistingClientKeys(IReadOnlyList<Client> clients)
    {
        var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var client in clients)
        {
            var key = ClientKey(client.AccountNumber, client.PppoeUsername, client.Name);
            if (!string.IsNullOrWhiteSpace(key))
            {
                keys.Add(key);
            }
        }

        return keys;
    }

    private static string ClientKey(string accountNumber, string pppoe, string name)
    {
        if (!string.IsNullOrWhiteSpace(pppoe))
        {
            return $"pppoe:{pppoe.Trim()}";
        }

        if (!string.IsNullOrWhiteSpace(accountNumber))
        {
            return $"account:{accountNumber.Trim()}";
        }

        return string.IsNullOrWhiteSpace(name) ? "" : $"name:{NormalizeName(name)}";
    }

    private static string NormalizeName(string value)
    {
        var withoutNotes = ParentheticalRegex().Replace(value.ToUpperInvariant(), "");
        return SpaceRegex().Replace(NonAlphaNumericRegex().Replace(withoutNotes, " "), " ").Trim();
    }

    private static object? Get(IReadOnlyList<object?> row, IReadOnlyDictionary<string, int> headers, string header)
    {
        return headers.TryGetValue(header, out var index) && index < row.Count ? row[index] : null;
    }

    private static object? GetAny(IReadOnlyList<object?> row, IReadOnlyDictionary<string, int> headers, params string[] headerNames)
    {
        foreach (var header in headerNames)
        {
            var value = Get(row, headers, header);
            if (!string.IsNullOrWhiteSpace(Text(value)))
            {
                return value;
            }
        }

        return null;
    }

    private static object? Cell(IReadOnlyList<object?> row, int? index)
    {
        return index is int value && value >= 0 && value < row.Count ? row[value] : null;
    }

    private static bool HasHeader(IEnumerable<object?> row, string header)
    {
        return row.Any(value => Text(value).Equals(header, StringComparison.OrdinalIgnoreCase));
    }

    private static Dictionary<string, int> HeaderIndexes(IReadOnlyList<object?> row)
    {
        var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < row.Count; index++)
        {
            var header = Text(row[index]).Trim();
            if (!string.IsNullOrWhiteSpace(header) && !headers.ContainsKey(header))
            {
                headers[header] = index;
            }
        }

        return headers;
    }

    private static Dictionary<string, List<int>> HeaderIndexesMulti(IReadOnlyList<object?> row)
    {
        var headers = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < row.Count; index++)
        {
            var header = Text(row[index]).Trim().ToUpperInvariant();
            if (string.IsNullOrWhiteSpace(header))
            {
                continue;
            }

            if (!headers.TryGetValue(header, out var indexes))
            {
                indexes = [];
                headers[header] = indexes;
            }

            indexes.Add(index);
        }

        return headers;
    }

    private static int? HeaderColumn(IReadOnlyDictionary<string, List<int>> headers, string header, int occurrence = 0)
    {
        return headers.TryGetValue(header, out var indexes) && indexes.Count > occurrence
            ? indexes[occurrence]
            : null;
    }

    private static string Text(object? value)
    {
        return value switch
        {
            null => "",
            double number when Math.Abs(number % 1) < 0.000001 => ((int)number).ToString(CultureInfo.InvariantCulture),
            IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture)?.Trim() ?? "",
            _ => value.ToString()?.Trim() ?? ""
        };
    }

    private static decimal Money(object? value)
    {
        if (value is double number)
        {
            return Convert.ToDecimal(number);
        }

        var text = Text(value).Replace(",", "", StringComparison.OrdinalIgnoreCase).Replace("PHP", "", StringComparison.OrdinalIgnoreCase);
        return decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out var amount) ? amount : 0;
    }

    private static decimal? PositiveMoney(object? value)
    {
        var amount = Money(value);
        return amount > 0 ? amount : null;
    }

    private static DateOnly? DateValue(object? value)
    {
        if (value is double serialNumber && serialNumber > 20000)
        {
            return DateOnly.FromDateTime(new DateTime(1899, 12, 30).AddDays(serialNumber));
        }

        if (value is int serialInteger && serialInteger > 20000)
        {
            return DateOnly.FromDateTime(new DateTime(1899, 12, 30).AddDays(serialInteger));
        }

        var text = Text(value);
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        if (DateOnly.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
        {
            return date;
        }

        foreach (var format in new[] { "yyyy-MM-dd", "M/d/yyyy", "MM/dd/yyyy", "MMM d, yyyy", "MMM dd, yyyy", "MMMM d, yyyy", "MMMM dd, yyyy" })
        {
            if (DateOnly.TryParseExact(text, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
            {
                return date;
            }
        }

        return null;
    }

    private static DateOnly? SheetMonth(string sheetName)
    {
        var match = SheetMonthRegex().Match(sheetName.ToUpperInvariant());
        if (!match.Success)
        {
            return null;
        }

        var monthName = match.Groups[1].Value;
        return MonthNumbers.TryGetValue(monthName, out var month)
            ? new DateOnly(int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture), month, 1)
            : null;
    }

    private static DateOnly MonthStart(DateOnly date)
    {
        return new DateOnly(date.Year, date.Month, 1);
    }

    private static Dictionary<string, List<List<object?>>> ReadWorkbook(Stream stream)
    {
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var sharedStrings = ReadSharedStrings(archive);
        var workbook = XDocument.Load(archive.GetEntry("xl/workbook.xml")!.Open());
        var relationships = XDocument.Load(archive.GetEntry("xl/_rels/workbook.xml.rels")!.Open());
        var targets = relationships.Root!
            .Elements(PackageRelationshipNs + "Relationship")
            .ToDictionary(element => element.Attribute("Id")!.Value, element => element.Attribute("Target")!.Value);

        var sheets = new Dictionary<string, List<List<object?>>>(StringComparer.OrdinalIgnoreCase);
        foreach (var sheet in workbook.Root!.Element(SpreadsheetNs + "sheets")!.Elements(SpreadsheetNs + "sheet"))
        {
            var name = sheet.Attribute("name")!.Value;
            var relationshipId = sheet.Attribute(RelationshipNs + "id")!.Value;
            var target = targets[relationshipId];
            var entryName = target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase)
                ? target
                : $"xl/{target.TrimStart('/')}";
            var sheetEntry = archive.GetEntry(entryName);
            if (sheetEntry is null)
            {
                continue;
            }

            sheets[name] = ReadSheet(sheetEntry, sharedStrings);
        }

        return sheets;
    }

    private static List<string> ReadSharedStrings(ZipArchive archive)
    {
        var entry = archive.GetEntry("xl/sharedStrings.xml");
        if (entry is null)
        {
            return [];
        }

        var document = XDocument.Load(entry.Open());
        return document.Root!
            .Elements(SpreadsheetNs + "si")
            .Select(item => string.Concat(item.Descendants(SpreadsheetNs + "t").Select(text => text.Value)))
            .ToList();
    }

    private static List<List<object?>> ReadSheet(ZipArchiveEntry sheetEntry, IReadOnlyList<string> sharedStrings)
    {
        var document = XDocument.Load(sheetEntry.Open());
        var rows = new List<List<object?>>();
        foreach (var row in document.Descendants(SpreadsheetNs + "row"))
        {
            var values = new List<object?>();
            foreach (var cell in row.Elements(SpreadsheetNs + "c"))
            {
                var cellReference = cell.Attribute("r")?.Value ?? "A1";
                var index = ColumnIndex(cellReference);
                while (values.Count <= index)
                {
                    values.Add(null);
                }

                values[index] = CellValue(cell, sharedStrings);
            }

            rows.Add(values);
        }

        return rows;
    }

    private static object? CellValue(XElement cell, IReadOnlyList<string> sharedStrings)
    {
        var type = cell.Attribute("t")?.Value;
        var value = cell.Element(SpreadsheetNs + "v")?.Value;
        if (type == "s" && int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var sharedIndex))
        {
            return sharedIndex >= 0 && sharedIndex < sharedStrings.Count ? sharedStrings[sharedIndex] : "";
        }

        if (type == "inlineStr")
        {
            return string.Concat(cell.Descendants(SpreadsheetNs + "t").Select(text => text.Value));
        }

        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
        {
            return number;
        }

        return value ?? "";
    }

    private static int ColumnIndex(string cellReference)
    {
        var total = 0;
        foreach (var character in cellReference.TakeWhile(char.IsLetter))
        {
            total = total * 26 + char.ToUpperInvariant(character) - 'A' + 1;
        }

        return total - 1;
    }

    [GeneratedRegex("[^A-Z0-9]+")]
    private static partial Regex NonAlphaNumericRegex();

    [GeneratedRegex(@"\([^)]*\)")]
    private static partial Regex ParentheticalRegex();

    [GeneratedRegex(@"\s+")]
    private static partial Regex SpaceRegex();

    [GeneratedRegex(@"^BILLS\s+[A-Z]{3,9}20\d{2}$", RegexOptions.IgnoreCase)]
    private static partial Regex BillsSheetRegex();

    [GeneratedRegex(@"^CASH\s+[A-Z]{3,9}20\d{2}$", RegexOptions.IgnoreCase)]
    private static partial Regex CashSheetRegex();

    [GeneratedRegex(@"^GCASH\s+[A-Z]{3,9}20\d{2}$", RegexOptions.IgnoreCase)]
    private static partial Regex GcashSheetRegex();

    [GeneratedRegex(@"\b([A-Z]+)\s*(20\d{2})\b")]
    private static partial Regex SheetMonthRegex();

    [GeneratedRegex(@"\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC|JANUARY|FEBRUARY|MARCH|APRIL)\b.*")]
    private static partial Regex PaymentMonthSuffixRegex();
}

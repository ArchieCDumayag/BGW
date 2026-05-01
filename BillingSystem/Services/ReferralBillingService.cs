using BillingSystem.Models;

namespace BillingSystem.Services;

public static class ReferralBillingService
{
    public static int ApplyReferralDiscounts(BillingData data, IEnumerable<Client> newClients)
    {
        var applied = 0;
        foreach (var client in newClients)
        {
            var before = data.Referrals.Count;
            ApplyReferralDiscount(data, client);
            if (data.Referrals.Count > before)
            {
                applied++;
            }
        }

        return applied;
    }

    public static string ApplyReferralDiscount(BillingData data, Client newClient)
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

    public static string NormalizeReferralText(string? referral)
    {
        return string.IsNullOrWhiteSpace(referral) ? "INQUIRE" : referral.Trim();
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

    private static string ReferralOptionText(Client client)
    {
        return $"{client.AccountNumber} - {client.Name}".Trim();
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

    private static decimal PlanAmountForBillingMonth(
        Client client,
        DateOnly month,
        IEnumerable<ClientPlanChange> planChanges)
    {
        var planChange = planChanges
            .Where(change => change.ClientId == client.Id && change.EffectiveMonth <= month)
            .OrderByDescending(change => change.EffectiveMonth)
            .ThenByDescending(change => change.Id)
            .FirstOrDefault();

        return planChange?.PlanAmount ?? client.PlanAmount;
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

    private static int NextId(IEnumerable<int> ids)
    {
        return ids.DefaultIfEmpty().Max() + 1;
    }
}

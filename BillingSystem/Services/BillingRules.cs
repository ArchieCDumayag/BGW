using BillingSystem.Models;

namespace BillingSystem.Services;

public static class BillingRules
{
    public const decimal XentronetEarlyDiscount = 200;

    public static string NormalizeBillingType(string? billingType)
    {
        var value = billingType?.Trim() ?? "";
        if (value.Equals("Postpaid", StringComparison.OrdinalIgnoreCase))
        {
            return "Postpaid";
        }

        if (value.Equals("Xentronet", StringComparison.OrdinalIgnoreCase)
            || value.Equals("XentroNet", StringComparison.OrdinalIgnoreCase))
        {
            return "Xentronet";
        }

        return "Prepaid";
    }

    public static decimal ProratedFirstBill(decimal planAmount, DateOnly installedOn, string? billingType)
    {
        if (planAmount <= 0)
        {
            return 0;
        }

        var type = NormalizeBillingType(billingType);
        var dueDate = FirstBillDueDate(installedOn, type);
        var cycleEnd = type.Equals("Postpaid", StringComparison.OrdinalIgnoreCase)
            ? dueDate
            : dueDate.Day == 1 && dueDate.Month == installedOn.Month
                ? dueDate.AddMonths(1).AddDays(-1)
                : dueDate.AddDays(-1);
        var daysInMonth = DateTime.DaysInMonth(installedOn.Year, installedOn.Month);
        var billableDays = Math.Max(0, cycleEnd.Day - installedOn.Day + 1);
        return Math.Round(planAmount / daysInMonth * billableDays, 2, MidpointRounding.AwayFromZero);
    }

    public static DateOnly FirstBillDueDate(DateOnly installedOn, string? billingType)
    {
        var type = NormalizeBillingType(billingType);
        if (type.Equals("Postpaid", StringComparison.OrdinalIgnoreCase))
        {
            return new DateOnly(installedOn.Year, installedOn.Month, DateTime.DaysInMonth(installedOn.Year, installedOn.Month));
        }

        var firstDay = new DateOnly(installedOn.Year, installedOn.Month, 1);
        return installedOn > firstDay ? firstDay.AddMonths(1) : firstDay;
    }

    public static DateOnly DueDateForBillingMonth(Client client, DateOnly month)
    {
        if (client.DateInstalled is DateOnly installed
            && installed.Year == month.Year
            && installed.Month == month.Month)
        {
            return FirstBillDueDate(installed, client.BillingType);
        }

        if (NormalizeBillingType(client.BillingType).Equals("Postpaid", StringComparison.OrdinalIgnoreCase))
        {
            return new DateOnly(month.Year, month.Month, DateTime.DaysInMonth(month.Year, month.Month));
        }

        return new DateOnly(month.Year, month.Month, 1);
    }

    public static ClientCollectionStatus CollectionStatusForClient(
        Client client,
        IEnumerable<Payment> payments,
        DateOnly? asOf = null)
    {
        var date = asOf ?? DateOnly.FromDateTime(DateTime.Today);
        var paidThisMonth = payments
            .Where(p => p.PaidOn.Year == date.Year && p.PaidOn.Month == date.Month)
            .Sum(p => p.Amount);
        var amountDue = CurrentAmountDue(client);
        var partialBalance = paidThisMonth > 0 && paidThisMonth < amountDue ? amountDue - paidThisMonth : 0;
        var unpaid = paidThisMonth == 0 ? amountDue : 0;
        var status = amountDue <= 0 && paidThisMonth > 0 ? "Paid" :
            amountDue <= 0 ? "No bill" :
            paidThisMonth >= amountDue ? "Paid" :
            paidThisMonth > 0 ? "Partial" : "Unpaid";

        return new ClientCollectionStatus
        {
            CurrentBill = client.Bills,
            AmountDue = amountDue,
            PaidThisMonth = paidThisMonth,
            Paid = paidThisMonth,
            Unpaid = unpaid,
            PartialBalance = partialBalance,
            Status = status
        };
    }

    public static BillingRuleInfo ForClient(Client client, DateOnly? asOf = null)
    {
        var date = asOf ?? DateOnly.FromDateTime(DateTime.Today);
        var billingType = NormalizeBillingType(client.BillingType);
        if (billingType.Equals("Postpaid", StringComparison.OrdinalIgnoreCase))
        {
            var dueDate = LastDayDue(date);
            return new BillingRuleInfo
            {
                BillingType = "Postpaid",
                ScheduleLabel = "Every last day of the month",
                NextDueDate = dueDate,
                Summary = $"Postpaid due every last day of the month. Next due: {dueDate:MMM dd, yyyy}."
            };
        }

        if (billingType.Equals("Xentronet", StringComparison.OrdinalIgnoreCase))
        {
            var dueDate = FirstDayDue(date);
            return new BillingRuleInfo
            {
                BillingType = "Xentronet",
                ScheduleLabel = "Prepaid, every 1st day of the month",
                NextDueDate = dueDate,
                HasEarlyDiscount = true,
                EarlyDiscountAmount = XentronetEarlyDiscount,
                DiscountDeadline = dueDate,
                Summary = $"Xentronet prepaid due every 1st day. PHP {XentronetEarlyDiscount:N0} discount available before {dueDate:MMM dd, yyyy} when admin applies it."
            };
        }

        var prepaidDueDate = FirstDayDue(date);
        return new BillingRuleInfo
        {
            BillingType = "Prepaid",
            ScheduleLabel = "Every 1st day of the month",
            NextDueDate = prepaidDueDate,
            Summary = $"Prepaid due every 1st day of the month. Next due: {prepaidDueDate:MMM dd, yyyy}."
        };
    }

    private static decimal CurrentAmountDue(Client client)
    {
        return client.Bills;
    }

    private static DateOnly FirstDayDue(DateOnly date)
    {
        var dueDate = new DateOnly(date.Year, date.Month, 1);
        return date <= dueDate ? dueDate : dueDate.AddMonths(1);
    }

    private static DateOnly LastDayDue(DateOnly date)
    {
        var dueDate = new DateOnly(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));
        if (date <= dueDate)
        {
            return dueDate;
        }

        var nextMonth = date.AddMonths(1);
        return new DateOnly(nextMonth.Year, nextMonth.Month, DateTime.DaysInMonth(nextMonth.Year, nextMonth.Month));
    }
}

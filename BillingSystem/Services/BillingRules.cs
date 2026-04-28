using BillingSystem.Models;

namespace BillingSystem.Services;

public static class BillingRules
{
    public const decimal XentronetEarlyDiscount = 200;

    public static ClientCollectionStatus CollectionStatusForClient(
        Client client,
        IEnumerable<Payment> payments,
        DateOnly? asOf = null)
    {
        var date = asOf ?? DateOnly.FromDateTime(DateTime.Today);
        var rule = ForClient(client, date);
        var paidThisMonth = payments
            .Where(p => p.PaidOn.Year == date.Year && p.PaidOn.Month == date.Month)
            .Sum(p => p.Amount);
        var amountDue = CurrentAmountDue(client, rule, date);
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
        if (client.BillingType.Equals("Postpaid", StringComparison.OrdinalIgnoreCase))
        {
            var dueDate = LastDayDue(date);
            return new BillingRuleInfo
            {
                BillingType = "Postpaid",
                ScheduleLabel = "Every last day of the month",
                NextDueDate = dueDate,
                DiscountedCurrentBill = client.Bills,
                Summary = $"Postpaid due every last day of the month. Next due: {dueDate:MMM dd, yyyy}."
            };
        }

        if (client.BillingType.Equals("Xentronet", StringComparison.OrdinalIgnoreCase))
        {
            var dueDate = FirstDayDue(date);
            var discountedBill = Math.Max(0, client.Bills - XentronetEarlyDiscount);
            return new BillingRuleInfo
            {
                BillingType = "Xentronet",
                ScheduleLabel = "Prepaid, every 1st day of the month",
                NextDueDate = dueDate,
                HasEarlyDiscount = true,
                EarlyDiscountAmount = XentronetEarlyDiscount,
                DiscountDeadline = dueDate,
                DiscountedCurrentBill = discountedBill,
                Summary = $"Xentronet prepaid due every 1st day. PHP {XentronetEarlyDiscount:N0} discount if paid before {dueDate:MMM dd, yyyy}."
            };
        }

        var prepaidDueDate = FirstDayDue(date);
        return new BillingRuleInfo
        {
            BillingType = "Prepaid",
            ScheduleLabel = "Every 1st day of the month",
            NextDueDate = prepaidDueDate,
            DiscountedCurrentBill = client.Bills,
            Summary = $"Prepaid due every 1st day of the month. Next due: {prepaidDueDate:MMM dd, yyyy}."
        };
    }

    private static decimal CurrentAmountDue(Client client, BillingRuleInfo rule, DateOnly date)
    {
        if (rule.HasEarlyDiscount && rule.DiscountDeadline is DateOnly deadline && date < deadline)
        {
            return rule.DiscountedCurrentBill;
        }

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

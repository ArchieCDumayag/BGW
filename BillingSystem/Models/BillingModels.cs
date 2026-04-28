using System.Text.Json.Serialization;

namespace BillingSystem.Models;

public sealed class BillingData
{
    public List<Client> Clients { get; set; } = [];
    public List<Payment> Payments { get; set; } = [];
    public List<Expense> Expenses { get; set; } = [];
    public List<Technician> Technicians { get; set; } = [];
    public List<JobTicket> Jobs { get; set; } = [];
    public List<PppoeUser> PppoeUsers { get; set; } = [];
    public List<TrafficSample> TrafficSamples { get; set; } = [];
    public List<UserAccount> UserAccounts { get; set; } = [];
    public List<ClientPlanChange> PlanChanges { get; set; } = [];
    public List<ClientMonthlyBillOverride> MonthlyBillOverrides { get; set; } = [];
    public List<ClientReferral> Referrals { get; set; } = [];
    public List<ServicePlan> Plans { get; set; } = [];
    public List<CoverageArea> CoverageAreas { get; set; } = [];
    public List<NapLocation> NapLocations { get; set; } = [];
    public List<OltDevice> OltDevices { get; set; } = [];
    public List<CollectorAssignment> CollectorAssignments { get; set; } = [];
    public List<SupportTicket> Tickets { get; set; } = [];
    public List<UserActivityLog> ActivityLogs { get; set; } = [];
    public SystemSettings Settings { get; set; } = new();
}

public sealed class Client
{
    public int Id { get; set; }
    public string AccountNumber { get; set; } = "";
    public DateOnly? DateInstalled { get; set; }
    public string Status { get; set; } = "Active";
    public string BillingType { get; set; } = "Prepaid";
    public decimal PlanAmount { get; set; }
    public string Area { get; set; } = "";
    public string Zone { get; set; } = "";
    public string Name { get; set; } = "";
    public string PppoeUsername { get; set; } = "";
    public string Contact { get; set; } = "";
    public string Email { get; set; } = "";
    public string FacebookAccount { get; set; } = "";
    public string Referral { get; set; } = "INQUIRE";
    public decimal Balance { get; set; }
    public decimal Advance { get; set; }
    public decimal Bills { get; set; }
    public string Address { get; set; } = "";
    public decimal? Latitude { get; set; }
    public decimal? Longitude { get; set; }
    public decimal CreditLimit { get; set; }
    public string Remarks { get; set; } = "";
}

public sealed class Payment
{
    public int Id { get; set; }
    public int ClientId { get; set; }
    public DateOnly PaidOn { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public decimal Amount { get; set; }
    public string Method { get; set; } = "Cash";
    public string ReferenceNumber { get; set; } = "";
    public string CollectedBy { get; set; } = "";
    public string Remarks { get; set; } = "";
}

public sealed class ClientPlanChange
{
    public int Id { get; set; }
    public int ClientId { get; set; }
    public DateOnly EffectiveMonth { get; set; }
    public decimal PlanAmount { get; set; }
    public DateTime RecordedAt { get; set; } = DateTime.Now;
    public string Remarks { get; set; } = "";
}

public sealed class ClientMonthlyBillOverride
{
    public int Id { get; set; }
    public int ClientId { get; set; }
    public DateOnly BillingMonth { get; set; }
    public decimal BillAmount { get; set; }
    public decimal? Balance { get; set; }
    public decimal? Advance { get; set; }
    public decimal DiscountAmount { get; set; }
    public string DiscountRemarks { get; set; } = "";
    public DateTime RecordedAt { get; set; } = DateTime.Now;
    public string Remarks { get; set; } = "";
}

public sealed class ClientReferral
{
    public int Id { get; set; }
    public int ReferrerClientId { get; set; }
    public int NewClientId { get; set; }
    public string ReferrerName { get; set; } = "";
    public string NewClientName { get; set; } = "";
    public string ReferralText { get; set; } = "";
    public DateOnly DiscountStartMonth { get; set; }
    public decimal DiscountAmount { get; set; }
    public decimal AppliedAmount { get; set; }
    public DateTime RecordedAt { get; set; } = DateTime.Now;
    public string Remarks { get; set; } = "";
}

public sealed class Expense
{
    public int Id { get; set; }
    public DateOnly SpentOn { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public string Category { get; set; } = "Operations";
    public string Description { get; set; } = "";
    public decimal Amount { get; set; }
}

public sealed class Technician
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public string Phone { get; set; } = "";
    public string Area { get; set; } = "";
    public bool IsActive { get; set; } = true;
}

public sealed class JobTicket
{
    public int Id { get; set; }
    public int ClientId { get; set; }
    public int? TechnicianId { get; set; }
    public string Type { get; set; } = "Repair";
    public string Status { get; set; } = "Open";
    public DateOnly ScheduledOn { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public string Remarks { get; set; } = "";
    public DateTime? CompletedAt { get; set; }
}

public sealed class PppoeUser
{
    public int Id { get; set; }
    public int ClientId { get; set; }
    public string Username { get; set; } = "";
    public string Status { get; set; } = "Unknown";
    public string IpAddress { get; set; } = "";
    public DateTime? LastSeenAt { get; set; }
}

public sealed class TrafficSample
{
    public int Id { get; set; }
    public int ClientId { get; set; }
    public DateTime CapturedAt { get; set; } = DateTime.Now;
    public decimal DownloadMbps { get; set; }
    public decimal UploadMbps { get; set; }
}

public sealed class UserAccount
{
    public int Id { get; set; }
    public string Username { get; set; } = "";
    public string Password { get; set; } = "";
    public string Role { get; set; } = "Technician";
    public string DisplayName { get; set; } = "";
    public int? TechnicianId { get; set; }
    public bool IsActive { get; set; } = true;
}

public sealed class UserActivityLog
{
    public int Id { get; set; }
    public DateTime OccurredAt { get; set; } = DateTime.Now;
    public int? UserId { get; set; }
    public string Username { get; set; } = "";
    public string DisplayName { get; set; } = "";
    public string Role { get; set; } = "";
    public string Action { get; set; } = "";
    public string Controller { get; set; } = "";
    public string Method { get; set; } = "";
    public string Path { get; set; } = "";
    public string IpAddress { get; set; } = "";
    public int StatusCode { get; set; }
    public string Details { get; set; } = "";
}

public sealed class ServicePlan
{
    public int Id { get; set; }
    public string PlanName { get; set; } = "";
    public decimal Price { get; set; }
    public string Type { get; set; } = "Prepaid";
}

public sealed class CoverageArea
{
    public int Id { get; set; }
    public string AreaName { get; set; } = "";
}

public sealed class NapLocation
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public string AreaName { get; set; } = "";
    public decimal? Latitude { get; set; }
    public decimal? Longitude { get; set; }
    public string Remarks { get; set; } = "";
}

public sealed class OltDevice
{
    public int Id { get; set; }
    public string OltName { get; set; } = "";
    public string Technology { get; set; } = "Gpon";
    public string Site { get; set; } = "";
    public int TotalPonPorts { get; set; }
}

public sealed class CollectorAssignment
{
    public int Id { get; set; }
    public int ClientId { get; set; }
    public int? CollectorUserId { get; set; }
    public string CollectorName { get; set; } = "";
    public string AreaName { get; set; } = "";
    public DateOnly AssignedOn { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public bool IsActive { get; set; } = true;
    public string Remarks { get; set; } = "";
}

public sealed class SupportTicket
{
    public int Id { get; set; }
    public int ClientId { get; set; }
    public string Subject { get; set; } = "";
    public string Type { get; set; } = "Repair";
    public string Priority { get; set; } = "Normal";
    public string Status { get; set; } = "Open";
    public DateOnly CreatedOn { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public int? AssignedTechnicianId { get; set; }
    public string AssignedTo { get; set; } = "";
    public string Remarks { get; set; } = "";
}

public sealed record TechnicianAssignmentOption(int Id, string Name);

public sealed class SystemSettings
{
    public string CompanyName { get; set; } = "Billing System";
    public string SystemDisplayName { get; set; } = "3JBGW Billing";
    public string CompanyLogoUrl { get; set; } = "";
    public string BrowserLogoUrl { get; set; } = "";
    public int MonthlyDueDay { get; set; } = 15;
    public string SmsReminderTemplate { get; set; } = "Hi {Name}, your balance is {Balance}.";
    public string Currency { get; set; } = "PHP";
    public string SemaphoreApiKey { get; set; } = "";
    public string SemaphoreSenderName { get; set; } = "";
    public string MikrotikHost { get; set; } = "";
    public int MikrotikApiPort { get; set; } = 8728;
    public string MikrotikApiUser { get; set; } = "";
    public string MikrotikApiPassword { get; set; } = "";
    public string GCashAccountName { get; set; } = "";
    public string GCashAccountNumber { get; set; } = "";
    public string GCashQrCodeUrl { get; set; } = "";
}

public sealed class DashboardViewModel
{
    public int TotalClients { get; init; }
    public int ActiveClients { get; init; }
    public int DisconnectedClients { get; init; }
    public decimal MonthlyRecurringRevenue { get; init; }
    public decimal TotalBalances { get; init; }
    public decimal PaymentsThisMonth { get; init; }
    public decimal ExpensesThisMonth { get; init; }
    public int OpenJobs { get; init; }
    public string CurrentMonthLabel { get; init; } = "";
    public decimal TotalBilled { get; init; }
    public decimal TotalCollected { get; init; }
    public decimal CollectionRate { get; init; }
    public int PayingSubscribers { get; init; }
    public decimal OutstandingBalance { get; init; }
    public decimal CashCollected { get; init; }
    public decimal GCashCollected { get; init; }
    public decimal OtherCollected { get; init; }
    public IReadOnlyList<DashboardMethodSummary> MethodSummaries { get; init; } = [];
    public IReadOnlyList<DashboardPaymentEntry> RecentPayments { get; init; } = [];
    public IReadOnlyList<Client> RecentClients { get; init; } = [];
    public IReadOnlyList<JobTicket> RecentJobs { get; init; } = [];
}

public sealed class DashboardMethodSummary
{
    public string Method { get; init; } = "";
    public decimal Amount { get; init; }
    public int Count { get; init; }
    public decimal Percentage { get; init; }
    public string BarClass { get; init; } = "bg-secondary";
    public string BadgeClass { get; init; } = "text-bg-secondary";
}

public sealed class DashboardPaymentEntry
{
    public DateOnly PaidOn { get; init; }
    public string ClientName { get; init; } = "";
    public string Method { get; init; } = "";
    public decimal Amount { get; init; }
    public string ReferenceNumber { get; init; } = "";
}

public sealed class ClientPaymentHistoryViewModel
{
    public Client Client { get; init; } = new();
    public PppoeUser Pppoe { get; init; } = new();
    public BillingRuleInfo BillingRule { get; init; } = new();
    public ClientCollectionStatus CollectionStatus { get; init; } = new();
    public IReadOnlyList<Payment> Payments { get; init; } = [];
    public decimal TotalPaid { get; init; }
}

public sealed class CustomerAccountViewModel
{
    public Client Client { get; init; } = new();
    public BillingRuleInfo BillingRule { get; init; } = new();
    public decimal CurrentBalance { get; init; }
    public IReadOnlyList<ClientPlanChange> PlanChanges { get; init; } = [];
    public IReadOnlyList<ClientMonthlyBillOverride> MonthlyBillOverrides { get; init; } = [];
    public IReadOnlyList<CustomerBillingMonth> BillingMonths { get; init; } = [];
    public IReadOnlyList<Payment> PaymentHistory { get; init; } = [];
    public int SelectedYear { get; init; }
    public IReadOnlyList<int> AvailableYears { get; init; } = [];
}

public sealed class CustomerBillingMonth
{
    public DateOnly Month { get; init; }
    public string MonthLabel { get; init; } = "";
    public DateOnly DueDate { get; init; }
    public decimal BillAmount { get; init; }
    public decimal AmountPaid { get; init; }
    public decimal Balance { get; init; }
    public decimal Advance { get; init; }
    public decimal DiscountAmount { get; init; }
    public string DiscountNote { get; init; } = "";
    public string Status { get; init; } = "";
    public string PaymentDates { get; init; } = "";
    public string Methods { get; init; } = "";
    public string References { get; init; } = "";
    public string Collectors { get; init; } = "";
    public string Remarks { get; init; } = "";
}

public sealed class PaymentReceiptViewModel
{
    public Payment Payment { get; init; } = new();
    public Client? Client { get; init; }
    public SystemSettings Settings { get; init; } = new();
    public BillingRuleInfo? BillingRule { get; init; }
    public decimal TotalPaid { get; init; }
    public string ReceiptNumber => $"OR-{Payment.Id:000000}";
}

public sealed class PppoeManagementViewModel
{
    public bool IsConnected { get; init; }
    public string ConnectionMessage { get; init; } = "";
    public string RouterHost { get; init; } = "";
    public string RouterIdentity { get; init; } = "";
    public string RouterAddress { get; init; } = "";
    public string Version { get; init; } = "";
    public string BoardName { get; init; } = "";
    public string Uptime { get; init; } = "";
    public int CpuLoad { get; init; }
    public long FreeMemory { get; init; }
    public long TotalMemory { get; init; }
    public int TotalUsers { get; init; }
    public int ActiveUsers { get; init; }
    public int OfflineUsers { get; init; }
    public int DisabledUsers { get; init; }
    public decimal TotalUsageGb { get; init; }
    public string Query { get; init; } = "";
    public string Filter { get; init; } = "All";
    public int Show { get; init; } = 25;
    public IReadOnlyList<PppoeAccountViewModel> Accounts { get; init; } = [];
}

public sealed record PppoeAccountViewModel
{
    public int Number { get; init; }
    public int? ClientId { get; init; }
    public string CustomerName { get; init; } = "";
    public string AccountNumber { get; init; } = "";
    public string Username { get; init; } = "";
    public string Address { get; init; } = "";
    public string CallerId { get; init; } = "";
    public string Profile { get; init; } = "";
    public string LastSeen { get; init; } = "";
    public string Status { get; init; } = "";
    public decimal UsageGb { get; init; }
    public bool IsAssigned { get; init; }
}

public sealed class BillingRuleInfo
{
    public string BillingType { get; init; } = "";
    public string ScheduleLabel { get; init; } = "";
    public DateOnly NextDueDate { get; init; }
    public bool HasEarlyDiscount { get; init; }
    public decimal EarlyDiscountAmount { get; init; }
    public DateOnly? DiscountDeadline { get; init; }
    public decimal DiscountedCurrentBill { get; init; }
    public string Summary { get; init; } = "";
}

public sealed class ClientCollectionStatus
{
    public decimal CurrentBill { get; init; }
    public decimal AmountDue { get; init; }
    public decimal PaidThisMonth { get; init; }
    public decimal Paid { get; init; }
    public decimal Unpaid { get; init; }
    public decimal PartialBalance { get; init; }
    public string Status { get; init; } = "";
}

public sealed record TechnicianRemarkRequest(string Remarks);
public sealed record CompleteJobRequest(string? Remarks);
public sealed record LoginViewModel(string Username, string Password, string Role, string? ReturnUrl = null, string? ErrorMessage = null);
public sealed record ApiLoginRequest(string Username, string Password, string Role);
public sealed record ApiLoginResponse(bool Success, string Message, string Username, string DisplayName, string Role, int? TechnicianId);

public sealed class TechnicianPortalViewModel
{
    public string DisplayName { get; init; } = "";
    public string Role { get; init; } = "";
    public int TechnicianId { get; init; }
    public IReadOnlyList<Client> AssignedClients { get; init; } = [];
    public IReadOnlyList<JobTicket> Jobs { get; init; } = [];
    public IReadOnlyList<SupportTicket> Tickets { get; init; } = [];
}

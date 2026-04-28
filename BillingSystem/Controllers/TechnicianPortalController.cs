using System.Security.Claims;
using BillingSystem.Models;
using BillingSystem.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace BillingSystem.Controllers;

[Authorize(Roles = "Technician,Collector,User")]
public sealed class TechnicianPortalController(IBillingStore store) : Controller
{
    public async Task<IActionResult> Index()
    {
        var data = await store.GetAsync();
        var role = User.FindFirstValue(ClaimTypes.Role) ?? "Technician";
        var technicianId = GetTechnicianId(role);
        var technician = data.Technicians.FirstOrDefault(t => t.Id == technicianId);
        var clients = technician is null ? Enumerable.Empty<Client>() : data.Clients.AsEnumerable();

        if (!string.IsNullOrWhiteSpace(technician?.Area) &&
            !technician.Area.Equals("All Areas", StringComparison.OrdinalIgnoreCase))
        {
            clients = clients.Where(c => c.Area.Equals(technician.Area, StringComparison.OrdinalIgnoreCase));
        }

        var isTechnician = role.Equals("Technician", StringComparison.OrdinalIgnoreCase);
        var jobs = isTechnician && technicianId > 0
            ? data.Jobs.Where(j => j.TechnicianId == technicianId || j.TechnicianId is null)
            : data.Jobs.Where(j => j.TechnicianId is null);

        var tickets = isTechnician && technicianId > 0
            ? data.Tickets.Where(ticket => ticket.AssignedTechnicianId == technicianId)
            : Enumerable.Empty<SupportTicket>();

        var model = new TechnicianPortalViewModel
        {
            DisplayName = User.FindFirstValue("DisplayName") ?? User.Identity?.Name ?? "Technician",
            Role = role,
            TechnicianId = technicianId,
            AssignedClients = clients.OrderBy(c => c.Area).ThenBy(c => c.Zone).ThenBy(c => c.Name).ToList(),
            Jobs = jobs
                .OrderBy(j => j.Status == "Done")
                .ThenBy(j => j.ScheduledOn)
                .ToList(),
            Tickets = tickets
                .OrderBy(ticket => ticket.Status.Equals("Closed", StringComparison.OrdinalIgnoreCase))
                .ThenByDescending(ticket => ticket.Priority.Equals("Urgent", StringComparison.OrdinalIgnoreCase))
                .ThenBy(ticket => ticket.CreatedOn)
                .ToList()
        };

        ViewBag.Clients = data.Clients.ToDictionary(c => c.Id);
        return View(model);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> AddRemark(int jobId, string remarks)
    {
        var data = await store.GetAsync();
        var job = data.Jobs.FirstOrDefault(j => j.Id == jobId);
        if (job is null)
        {
            return NotFound();
        }

        var stamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
        job.Remarks = string.IsNullOrWhiteSpace(job.Remarks)
            ? $"[{stamp}] {remarks}"
            : $"{job.Remarks}{Environment.NewLine}[{stamp}] {remarks}";

        await store.SaveAsync(data);
        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> MarkDone(int jobId, string? remarks)
    {
        var data = await store.GetAsync();
        var job = data.Jobs.FirstOrDefault(j => j.Id == jobId);
        if (job is null)
        {
            return NotFound();
        }

        job.Status = "Done";
        job.CompletedAt = DateTime.Now;
        if (!string.IsNullOrWhiteSpace(remarks))
        {
            job.Remarks = string.IsNullOrWhiteSpace(job.Remarks)
                ? remarks
                : $"{job.Remarks}{Environment.NewLine}{remarks}";
        }

        await store.SaveAsync(data);
        return RedirectToAction(nameof(Index));
    }

    private int GetTechnicianId(string role)
    {
        if (int.TryParse(User.FindFirstValue("TechnicianId"), out var technicianId) && technicianId > 0)
        {
            return technicianId;
        }

        return role.Equals("Technician", StringComparison.OrdinalIgnoreCase) &&
            int.TryParse(User.FindFirstValue(ClaimTypes.NameIdentifier), out var userId)
            ? userId
            : 0;
    }
}

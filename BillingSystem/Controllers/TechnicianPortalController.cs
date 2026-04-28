using System.Security.Claims;
using BillingSystem.Models;
using BillingSystem.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace BillingSystem.Controllers;

[Authorize(Roles = "Technician,Collector")]
public sealed class TechnicianPortalController(IBillingStore store) : Controller
{
    public async Task<IActionResult> Index()
    {
        var data = await store.GetAsync();
        var technicianId = GetTechnicianId();
        var technician = data.Technicians.FirstOrDefault(t => t.Id == technicianId);
        var clients = data.Clients.AsEnumerable();

        if (!string.IsNullOrWhiteSpace(technician?.Area) &&
            !technician.Area.Equals("All Areas", StringComparison.OrdinalIgnoreCase))
        {
            clients = clients.Where(c => c.Area.Equals(technician.Area, StringComparison.OrdinalIgnoreCase));
        }

        var model = new TechnicianPortalViewModel
        {
            DisplayName = User.FindFirstValue("DisplayName") ?? User.Identity?.Name ?? "Technician",
            Role = User.FindFirstValue(ClaimTypes.Role) ?? "Technician",
            TechnicianId = technicianId,
            AssignedClients = clients.OrderBy(c => c.Area).ThenBy(c => c.Zone).ThenBy(c => c.Name).ToList(),
            Jobs = data.Jobs
                .Where(j => j.TechnicianId == technicianId || j.TechnicianId is null)
                .OrderBy(j => j.Status == "Done")
                .ThenBy(j => j.ScheduledOn)
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

    private int GetTechnicianId()
    {
        return int.TryParse(User.FindFirstValue("TechnicianId"), out var technicianId)
            ? technicianId
            : 1;
    }
}

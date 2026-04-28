using BillingSystem.Models;
using BillingSystem.Services;
using Microsoft.AspNetCore.Mvc;

namespace BillingSystem.Controllers;

[ApiController]
[Route("api/technician")]
public sealed class TechnicianApiController(IBillingStore store) : ControllerBase
{
    [HttpGet("clients")]
    public async Task<IActionResult> AssignedClients([FromQuery] int technicianId = 1, [FromQuery] string? area = null)
    {
        var data = await store.GetAsync();
        var technician = data.Technicians.FirstOrDefault(t => t.Id == technicianId);
        var targetArea = area ?? technician?.Area;

        var clients = data.Clients.AsEnumerable();
        if (!string.IsNullOrWhiteSpace(targetArea) && !targetArea.Equals("All Areas", StringComparison.OrdinalIgnoreCase))
        {
            clients = clients.Where(c => c.Area.Equals(targetArea, StringComparison.OrdinalIgnoreCase));
        }

        return Ok(clients.OrderBy(c => c.Area).ThenBy(c => c.Zone).ThenBy(c => c.Name));
    }

    [HttpGet("jobs")]
    public async Task<IActionResult> Jobs([FromQuery] int technicianId = 1)
    {
        var data = await store.GetAsync();
        var jobs = data.Jobs
            .Where(j => j.TechnicianId == technicianId || j.TechnicianId is null)
            .OrderBy(j => j.Status == "Done")
            .ThenBy(j => j.ScheduledOn)
            .Select(j => new
            {
                Job = j,
                Client = data.Clients.FirstOrDefault(c => c.Id == j.ClientId)
            });

        return Ok(jobs);
    }

    [HttpGet("clients/{id:int}")]
    public async Task<IActionResult> ClientDetails(int id)
    {
        var data = await store.GetAsync();
        var client = data.Clients.FirstOrDefault(c => c.Id == id);
        if (client is null)
        {
            return NotFound();
        }

        return Ok(new
        {
            Client = client,
            BillingRule = BillingRules.ForClient(client),
            Pppoe = data.PppoeUsers.FirstOrDefault(p => p.ClientId == id) ?? new PppoeUser
            {
                ClientId = id,
                Username = client.PppoeUsername,
                Status = client.Status
            },
            OpenJobs = data.Jobs.Where(j => j.ClientId == id && j.Status != "Done").OrderBy(j => j.ScheduledOn),
            Payments = data.Payments.Where(p => p.ClientId == id).OrderByDescending(p => p.PaidOn).Take(12)
        });
    }

    [HttpPost("jobs/{id:int}/remarks")]
    public async Task<IActionResult> AddRemarks(int id, TechnicianRemarkRequest request)
    {
        var data = await store.GetAsync();
        var job = data.Jobs.FirstOrDefault(j => j.Id == id);
        if (job is null)
        {
            return NotFound();
        }

        var stamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
        job.Remarks = string.IsNullOrWhiteSpace(job.Remarks)
            ? $"[{stamp}] {request.Remarks}"
            : $"{job.Remarks}{Environment.NewLine}[{stamp}] {request.Remarks}";

        await store.SaveAsync(data);
        return Ok(job);
    }

    [HttpPost("jobs/{id:int}/done")]
    public async Task<IActionResult> MarkDone(int id, CompleteJobRequest request)
    {
        var data = await store.GetAsync();
        var job = data.Jobs.FirstOrDefault(j => j.Id == id);
        if (job is null)
        {
            return NotFound();
        }

        job.Status = "Done";
        job.CompletedAt = DateTime.Now;
        if (!string.IsNullOrWhiteSpace(request.Remarks))
        {
            job.Remarks = string.IsNullOrWhiteSpace(job.Remarks)
                ? request.Remarks
                : $"{job.Remarks}{Environment.NewLine}{request.Remarks}";
        }

        await store.SaveAsync(data);
        return Ok(job);
    }

    [HttpGet("clients/{id:int}/pppoe")]
    public async Task<IActionResult> PppoeStatus(int id)
    {
        var data = await store.GetAsync();
        var client = data.Clients.FirstOrDefault(c => c.Id == id);
        if (client is null)
        {
            return NotFound();
        }

        return Ok(data.PppoeUsers.FirstOrDefault(p => p.ClientId == id) ?? new PppoeUser
        {
            ClientId = id,
            Username = client.PppoeUsername,
            Status = client.Status
        });
    }
}

using System.Net;
using System.Text.RegularExpressions;
using BillingSystem.Models;

namespace BillingSystem.Services;

public interface IOltWebClient
{
    Task<OltSyncResult> GetOnuClientsAsync(OltDevice olt, CancellationToken cancellationToken = default);
}

public sealed class OltWebClient : IOltWebClient
{
    private static readonly string[] DefaultAuthPageCandidates =
    [
        "gpononuauthinfo.html",
        "onuauthinfo.html",
        "epononuauthinfo.html"
    ];

    private static readonly Regex AuthPageLinkRegex = new(
        "href\\s*=\\s*[\"'](?<href>[^\"']*onuauthinfo\\.html[^\"']*)[\"']",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex PonOptionRegex = new(
        "<option\\s+[^>]*pon\\s*=\\s*[\"']?(?<pon>\\d+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex OnuRowRegex = new(
        "<tr>\\s*<td>\\s*(?<onu>GPON[^<]+)</td>\\s*<td>(?<status>.*?)</td>\\s*<td>(?<description>.*?)</td>\\s*<td>(?<model>.*?)</td>\\s*<td>(?<profile>.*?)</td>\\s*<td>(?<mode>.*?)</td>\\s*<td>(?<info>.*?)</td>",
        RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);

    private static readonly Regex TagRegex = new("<.*?>", RegexOptions.Singleline | RegexOptions.Compiled);
    private static readonly Regex WhitespaceRegex = new("\\s+", RegexOptions.Compiled);

    public async Task<OltSyncResult> GetOnuClientsAsync(OltDevice olt, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(olt.ManagementUrl))
        {
            return OltSyncResult.Failed(olt, "Management URL is missing.");
        }

        if (string.IsNullOrWhiteSpace(olt.Username) || string.IsNullOrWhiteSpace(olt.Password))
        {
            return OltSyncResult.Failed(olt, "OLT username or password is missing.");
        }

        try
        {
            using var handler = new HttpClientHandler
            {
                AllowAutoRedirect = true,
                CookieContainer = new CookieContainer(),
                ServerCertificateCustomValidationCallback = HttpClientHandler.DangerousAcceptAnyServerCertificateValidator
            };
            using var http = new HttpClient(handler)
            {
                BaseAddress = OltActionBaseUri(olt.ManagementUrl),
                Timeout = TimeSpan.FromSeconds(18)
            };
            http.DefaultRequestHeaders.UserAgent.ParseAdd("BillingSystem/1.0");

            var login = await http.PostAsync(
                "main.html",
                new FormUrlEncodedContent(new Dictionary<string, string>
                {
                    ["user"] = olt.Username,
                    ["pass"] = olt.Password,
                    ["who"] = "100"
                }),
                cancellationToken);
            var loginHtml = await login.Content.ReadAsStringAsync(cancellationToken);
            if (!login.IsSuccessStatusCode || loginHtml.Contains("Sorry, you do not have access", StringComparison.OrdinalIgnoreCase))
            {
                return OltSyncResult.Failed(olt, "OLT login failed.");
            }

            var authPage = await GetFirstAvailableAuthPageAsync(http, loginHtml, cancellationToken);
            if (authPage is null)
            {
                return OltSyncResult.Failed(olt, "ONU authentication page was not found on this OLT.");
            }

            var firstPage = authPage.Value.Html;
            var authPagePath = authPage.Value.Path;
            var ponPorts = ParsePonPorts(firstPage);
            if (ponPorts.Count == 0)
            {
                ponPorts = olt.TotalPonPorts > 0
                    ? Enumerable.Range(1, Math.Min(olt.TotalPonPorts, 32)).ToList()
                    : [1];
            }

            var clients = new List<OltOnuClient>();
            foreach (var pon in ponPorts)
            {
                var page = pon == ponPorts[0]
                    ? firstPage
                    : await http.GetStringAsync($"{authPagePath}?slotid=0&portid={pon}&pon_select={pon}", cancellationToken);

                clients.AddRange(ParseAuthRows(olt, page));
            }

            return OltSyncResult.Succeeded(olt, ponPorts.Count, clients);
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or OperationCanceledException or UriFormatException)
        {
            return OltSyncResult.Failed(olt, ex.Message);
        }
    }

    private static async Task<(string Path, string Html)?> GetFirstAvailableAuthPageAsync(
        HttpClient http,
        string menuHtml,
        CancellationToken cancellationToken)
    {
        foreach (var candidate in AuthPageCandidates(menuHtml))
        {
            using var response = await http.GetAsync(candidate, cancellationToken);
            if (!response.IsSuccessStatusCode)
            {
                continue;
            }

            return (candidate, await response.Content.ReadAsStringAsync(cancellationToken));
        }

        return null;
    }

    private static IEnumerable<string> AuthPageCandidates(string menuHtml)
    {
        return AuthPageLinkRegex.Matches(menuHtml)
            .Select(match => NormalizePagePath(match.Groups["href"].Value))
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Concat(DefaultAuthPageCandidates)
            .Distinct(StringComparer.OrdinalIgnoreCase);
    }

    private static string NormalizePagePath(string href)
    {
        var value = href.Trim();
        if (Uri.TryCreate(value, UriKind.Absolute, out var absoluteUri))
        {
            value = absoluteUri.AbsolutePath;
        }

        var pathOnly = value.Split('?', '#')[0].Replace('\\', '/');
        var slashIndex = pathOnly.LastIndexOf('/');
        return slashIndex >= 0 ? pathOnly[(slashIndex + 1)..] : pathOnly;
    }

    private static Uri OltActionBaseUri(string managementUrl)
    {
        var uri = new Uri(managementUrl, UriKind.Absolute);
        var path = uri.AbsolutePath;
        var actionIndex = path.IndexOf("/action/", StringComparison.OrdinalIgnoreCase);
        var basePath = actionIndex >= 0
            ? path[..(actionIndex + "/action/".Length)]
            : "/";

        return new UriBuilder(uri.Scheme, uri.Host, uri.Port, basePath).Uri;
    }

    private static List<int> ParsePonPorts(string html)
    {
        return PonOptionRegex.Matches(html)
            .Select(match => int.TryParse(match.Groups["pon"].Value, out var port) ? port : 0)
            .Where(port => port > 0)
            .Distinct()
            .Order()
            .ToList();
    }

    private static IEnumerable<OltOnuClient> ParseAuthRows(OltDevice olt, string html)
    {
        foreach (Match match in OnuRowRegex.Matches(html))
        {
            var onuId = CleanHtml(match.Groups["onu"].Value);
            if (string.IsNullOrWhiteSpace(onuId))
            {
                continue;
            }

            yield return new OltOnuClient
            {
                OltDeviceId = olt.Id,
                OltName = olt.OltName,
                PonPort = ParsePonPort(onuId),
                OnuId = onuId,
                Status = CleanHtml(match.Groups["status"].Value),
                Description = CleanHtml(match.Groups["description"].Value),
                Model = CleanHtml(match.Groups["model"].Value),
                Profile = CleanHtml(match.Groups["profile"].Value),
                AuthMode = CleanHtml(match.Groups["mode"].Value),
                SerialNumber = CleanHtml(match.Groups["info"].Value),
                SyncedAt = DateTime.Now
            };
        }
    }

    private static string ParsePonPort(string onuId)
    {
        var match = Regex.Match(onuId, @"GPON\d+/(?<pon>\d+):", RegexOptions.IgnoreCase);
        return match.Success ? $"PON{match.Groups["pon"].Value}" : "";
    }

    private static string CleanHtml(string value)
    {
        var withoutTags = TagRegex.Replace(value, "");
        var decoded = WebUtility.HtmlDecode(withoutTags);
        return WhitespaceRegex.Replace(decoded, " ").Trim();
    }
}

public sealed record OltSyncResult(
    int OltDeviceId,
    string OltName,
    bool IsSuccess,
    int PonPortCount,
    IReadOnlyList<OltOnuClient> Clients,
    string ErrorMessage)
{
    public static OltSyncResult Succeeded(OltDevice olt, int ponPortCount, IReadOnlyList<OltOnuClient> clients)
    {
        return new OltSyncResult(olt.Id, olt.OltName, true, ponPortCount, clients, "");
    }

    public static OltSyncResult Failed(OltDevice olt, string errorMessage)
    {
        return new OltSyncResult(olt.Id, olt.OltName, false, 0, [], errorMessage);
    }
}

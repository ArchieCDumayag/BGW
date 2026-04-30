using System.Net.Sockets;
using System.Text;

namespace BillingSystem.Services;

public sealed class MikrotikRouterOsClient(string host, int port, string username, string password)
{
    public async Task<MikrotikSnapshot> GetSnapshotAsync(CancellationToken cancellationToken = default)
    {
        using var tcp = new TcpClient();
        await tcp.ConnectAsync(host, port, cancellationToken);
        await using var stream = tcp.GetStream();

        await LoginAsync(stream, cancellationToken);

        var identity = (await QueryAsync(stream, "/system/identity/print", [], cancellationToken)).FirstOrDefault() ?? [];
        var resource = (await QueryAsync(stream, "/system/resource/print", [], cancellationToken)).FirstOrDefault() ?? [];
        var board = (await QueryAsync(stream, "/system/routerboard/print", [], cancellationToken)).FirstOrDefault() ?? [];
        var addresses = await QueryAsync(stream, "/ip/address/print", ["=.proplist=address,interface,disabled"], cancellationToken);
        var secrets = await QueryAsync(stream, "/ppp/secret/print", ["=.proplist=.id,name,profile,caller-id,remote-address,disabled,comment"], cancellationToken);
        var profiles = await QueryOptionalAsync(stream, "/ppp/profile/print", ["=.proplist=.id,name,local-address,remote-address,rate-limit,only-one,dns-server,comment"], cancellationToken);
        var active = await QueryAsync(stream, "/ppp/active/print", ["=.proplist=name,address,caller-id,uptime,bytes-in,bytes-out"], cancellationToken);

        return new MikrotikSnapshot
        {
            Identity = Value(identity, "name"),
            Address = PrimaryAddress(addresses),
            Version = Value(resource, "version"),
            BoardName = FirstNonEmpty(Value(board, "model"), Value(resource, "board-name")),
            Uptime = Value(resource, "uptime"),
            CpuLoad = ToInt(Value(resource, "cpu-load")),
            FreeMemory = ToLong(Value(resource, "free-memory")),
            TotalMemory = ToLong(Value(resource, "total-memory")),
            Secrets = secrets.Select(ToSecret).Where(s => !string.IsNullOrWhiteSpace(s.Name)).ToList(),
            Profiles = profiles.Select(ToProfile).Where(p => !string.IsNullOrWhiteSpace(p.Name)).ToList(),
            ActiveSessions = active.Select(ToActiveSession).Where(s => !string.IsNullOrWhiteSpace(s.Name)).ToList()
        };
    }

    private async Task LoginAsync(NetworkStream stream, CancellationToken cancellationToken)
    {
        await WriteSentenceAsync(stream, ["/login", $"=name={username}", $"=password={password}"], cancellationToken);
        while (true)
        {
            var sentence = await ReadSentenceAsync(stream, cancellationToken);
            if (sentence.Count == 0)
            {
                continue;
            }

            if (sentence[0] == "!done")
            {
                return;
            }

            if (sentence[0] == "!trap")
            {
                throw new InvalidOperationException(TrapMessage(sentence, "MikroTik login failed."));
            }
        }
    }

    private static async Task<List<Dictionary<string, string>>> QueryAsync(
        NetworkStream stream,
        string command,
        IReadOnlyList<string> arguments,
        CancellationToken cancellationToken)
    {
        await WriteSentenceAsync(stream, [command, .. arguments], cancellationToken);

        var rows = new List<Dictionary<string, string>>();
        string? trapMessage = null;
        while (true)
        {
            var sentence = await ReadSentenceAsync(stream, cancellationToken);
            if (sentence.Count == 0)
            {
                continue;
            }

            if (sentence[0] == "!done")
            {
                if (!string.IsNullOrWhiteSpace(trapMessage))
                {
                    throw new InvalidOperationException(trapMessage);
                }

                return rows;
            }

            if (sentence[0] == "!trap")
            {
                trapMessage = TrapMessage(sentence, $"MikroTik command failed: {command}");
                continue;
            }

            if (sentence[0] == "!re")
            {
                rows.Add(ParseAttributes(sentence));
            }
        }
    }

    private static async Task<List<Dictionary<string, string>>> QueryOptionalAsync(
        NetworkStream stream,
        string command,
        IReadOnlyList<string> arguments,
        CancellationToken cancellationToken)
    {
        try
        {
            return await QueryAsync(stream, command, arguments, cancellationToken);
        }
        catch (InvalidOperationException)
        {
            return [];
        }
    }

    private static MikrotikPppoeSecret ToSecret(Dictionary<string, string> row)
    {
        return new MikrotikPppoeSecret
        {
            Id = Value(row, ".id"),
            Name = Value(row, "name"),
            Profile = Value(row, "profile"),
            CallerId = Value(row, "caller-id"),
            RemoteAddress = Value(row, "remote-address"),
            Disabled = IsTruthy(Value(row, "disabled")),
            Comment = Value(row, "comment")
        };
    }

    private static MikrotikPppoeActiveSession ToActiveSession(Dictionary<string, string> row)
    {
        return new MikrotikPppoeActiveSession
        {
            Name = Value(row, "name"),
            Address = Value(row, "address"),
            CallerId = Value(row, "caller-id"),
            Uptime = Value(row, "uptime"),
            BytesIn = ToLong(Value(row, "bytes-in")),
            BytesOut = ToLong(Value(row, "bytes-out"))
        };
    }

    private static MikrotikPppoeProfile ToProfile(Dictionary<string, string> row)
    {
        return new MikrotikPppoeProfile
        {
            Id = Value(row, ".id"),
            Name = Value(row, "name"),
            LocalAddress = Value(row, "local-address"),
            RemoteAddress = Value(row, "remote-address"),
            RateLimit = Value(row, "rate-limit"),
            OnlyOne = Value(row, "only-one"),
            DnsServer = Value(row, "dns-server"),
            Comment = Value(row, "comment")
        };
    }

    private string PrimaryAddress(IEnumerable<Dictionary<string, string>> addresses)
    {
        foreach (var address in addresses)
        {
            if (IsTruthy(Value(address, "disabled")))
            {
                continue;
            }

            var value = Value(address, "address");
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value.Split('/')[0];
            }
        }

        return host;
    }

    private static Dictionary<string, string> ParseAttributes(IReadOnlyList<string> sentence)
    {
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var word in sentence.Skip(1))
        {
            if (!word.StartsWith('='))
            {
                continue;
            }

            var secondEquals = word.IndexOf('=', 1);
            if (secondEquals <= 1)
            {
                continue;
            }

            values[word[1..secondEquals]] = word[(secondEquals + 1)..];
        }

        return values;
    }

    private static async Task WriteSentenceAsync(NetworkStream stream, IEnumerable<string> words, CancellationToken cancellationToken)
    {
        foreach (var word in words)
        {
            var bytes = Encoding.UTF8.GetBytes(word);
            await WriteLengthAsync(stream, bytes.Length, cancellationToken);
            await stream.WriteAsync(bytes, cancellationToken);
        }

        await WriteLengthAsync(stream, 0, cancellationToken);
        await stream.FlushAsync(cancellationToken);
    }

    private static async Task<List<string>> ReadSentenceAsync(NetworkStream stream, CancellationToken cancellationToken)
    {
        var words = new List<string>();
        while (true)
        {
            var length = await ReadLengthAsync(stream, cancellationToken);
            if (length == 0)
            {
                return words;
            }

            var buffer = new byte[length];
            var offset = 0;
            while (offset < buffer.Length)
            {
                var read = await stream.ReadAsync(buffer.AsMemory(offset, buffer.Length - offset), cancellationToken);
                if (read == 0)
                {
                    throw new IOException("MikroTik closed the API connection.");
                }

                offset += read;
            }

            words.Add(Encoding.UTF8.GetString(buffer));
        }
    }

    private static async Task WriteLengthAsync(NetworkStream stream, int length, CancellationToken cancellationToken)
    {
        byte[] bytes;
        if (length < 0x80)
        {
            bytes = [(byte)length];
        }
        else if (length < 0x4000)
        {
            bytes = [(byte)((length >> 8) | 0x80), (byte)(length & 0xFF)];
        }
        else if (length < 0x200000)
        {
            bytes = [(byte)((length >> 16) | 0xC0), (byte)((length >> 8) & 0xFF), (byte)(length & 0xFF)];
        }
        else
        {
            bytes =
            [
                (byte)((length >> 24) | 0xE0),
                (byte)((length >> 16) & 0xFF),
                (byte)((length >> 8) & 0xFF),
                (byte)(length & 0xFF)
            ];
        }

        await stream.WriteAsync(bytes, cancellationToken);
    }

    private static async Task<int> ReadLengthAsync(NetworkStream stream, CancellationToken cancellationToken)
    {
        var first = await ReadByteAsync(stream, cancellationToken);
        if ((first & 0x80) == 0)
        {
            return first;
        }

        if ((first & 0xC0) == 0x80)
        {
            return ((first & ~0xC0) << 8) + await ReadByteAsync(stream, cancellationToken);
        }

        if ((first & 0xE0) == 0xC0)
        {
            return ((first & ~0xE0) << 16)
                + (await ReadByteAsync(stream, cancellationToken) << 8)
                + await ReadByteAsync(stream, cancellationToken);
        }

        return ((first & ~0xF0) << 24)
            + (await ReadByteAsync(stream, cancellationToken) << 16)
            + (await ReadByteAsync(stream, cancellationToken) << 8)
            + await ReadByteAsync(stream, cancellationToken);
    }

    private static async Task<int> ReadByteAsync(NetworkStream stream, CancellationToken cancellationToken)
    {
        var buffer = new byte[1];
        var read = await stream.ReadAsync(buffer, cancellationToken);
        if (read != 1)
        {
            throw new IOException("MikroTik closed the API connection.");
        }

        return buffer[0];
    }

    private static string Value(IReadOnlyDictionary<string, string> values, string key)
    {
        return values.TryGetValue(key, out var value) ? value : "";
    }

    private static string FirstNonEmpty(params string[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
    }

    private static long ToLong(string value)
    {
        return long.TryParse(value, out var result) ? result : 0;
    }

    private static int ToInt(string value)
    {
        return int.TryParse(value, out var result) ? result : 0;
    }

    private static bool IsTruthy(string value)
    {
        return value.Equals("true", StringComparison.OrdinalIgnoreCase)
            || value.Equals("yes", StringComparison.OrdinalIgnoreCase);
    }

    private static string TrapMessage(IReadOnlyList<string> sentence, string fallback)
    {
        var values = ParseAttributes(sentence);
        return values.TryGetValue("message", out var message) && !string.IsNullOrWhiteSpace(message)
            ? message
            : fallback;
    }
}

public sealed class MikrotikSnapshot
{
    public string Identity { get; set; } = "";
    public string Address { get; set; } = "";
    public string Version { get; set; } = "";
    public string BoardName { get; set; } = "";
    public string Uptime { get; set; } = "";
    public int CpuLoad { get; set; }
    public long FreeMemory { get; set; }
    public long TotalMemory { get; set; }
    public IReadOnlyList<MikrotikPppoeSecret> Secrets { get; set; } = [];
    public IReadOnlyList<MikrotikPppoeProfile> Profiles { get; set; } = [];
    public IReadOnlyList<MikrotikPppoeActiveSession> ActiveSessions { get; set; } = [];
}

public sealed class MikrotikPppoeSecret
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string Profile { get; set; } = "";
    public string CallerId { get; set; } = "";
    public string RemoteAddress { get; set; } = "";
    public bool Disabled { get; set; }
    public string Comment { get; set; } = "";
}

public sealed class MikrotikPppoeActiveSession
{
    public string Name { get; set; } = "";
    public string Address { get; set; } = "";
    public string CallerId { get; set; } = "";
    public string Uptime { get; set; } = "";
    public long BytesIn { get; set; }
    public long BytesOut { get; set; }
}

public sealed class MikrotikPppoeProfile
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string LocalAddress { get; set; } = "";
    public string RemoteAddress { get; set; } = "";
    public string RateLimit { get; set; } = "";
    public string OnlyOne { get; set; } = "";
    public string DnsServer { get; set; } = "";
    public string Comment { get; set; } = "";
}

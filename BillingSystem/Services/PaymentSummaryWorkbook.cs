using System.Globalization;
using System.IO.Compression;
using System.Xml;
using BillingSystem.Models;

namespace BillingSystem.Services;

public static class PaymentSummaryWorkbook
{
    private const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";

    public static byte[] Create(BillingData data)
    {
        var payments = data.Payments
            .OrderByDescending(p => p.PaidOn)
            .ThenByDescending(p => p.Id)
            .ToList();
        var clients = data.Clients.ToDictionary(c => c.Id);

        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteText(archive, "[Content_Types].xml", ContentTypesXml());
            WriteText(archive, "_rels/.rels", PackageRelationshipsXml());
            WriteText(archive, "xl/workbook.xml", WorkbookXml());
            WriteText(archive, "xl/_rels/workbook.xml.rels", WorkbookRelationshipsXml());
            WriteText(archive, "xl/styles.xml", StylesXml());
            WriteWorksheet(archive, "xl/worksheets/sheet1.xml", BuildSummaryRows(payments), [28, 18, 18, 18]);
            WriteWorksheet(archive, "xl/worksheets/sheet2.xml", BuildPaymentRows(payments, clients), [14, 28, 14, 32, 18, 16, 16, 14, 22, 18, 36], freezeHeader: true);
        }

        return stream.ToArray();
    }

    private static List<object?[]> BuildSummaryRows(IReadOnlyList<Payment> payments)
    {
        var rows = new List<object?[]>
        {
            new object?[] { "Payments Summary" },
            new object?[] { "Generated", DateTime.Now.ToString("yyyy-MM-dd HH:mm") },
            Array.Empty<object?>(),
            new object?[] { "Total records", payments.Count },
            new object?[] { "Total collected", payments.Sum(p => p.Amount) },
            Array.Empty<object?>(),
            new object?[] { "By method" },
            new object?[] { "Method", "Count", "Amount" }
        };

        rows.AddRange(payments
            .GroupBy(p => NormalizePaymentMethod(p.Method))
            .OrderBy(g => g.Key)
            .Select(g => new object?[] { g.Key, g.Count(), g.Sum(p => p.Amount) }));

        rows.Add(Array.Empty<object?>());
        rows.Add(new object?[] { "By month" });
        rows.Add(new object?[] { "Month", "Count", "Amount" });
        rows.AddRange(payments
            .GroupBy(p => new DateOnly(p.PaidOn.Year, p.PaidOn.Month, 1))
            .OrderByDescending(g => g.Key)
            .Select(g => new object?[] { g.Key.ToString("MMMM yyyy"), g.Count(), g.Sum(p => p.Amount) }));

        return rows;
    }

    private static List<object?[]> BuildPaymentRows(IReadOnlyList<Payment> payments, IReadOnlyDictionary<int, Client> clients)
    {
        var rows = new List<object?[]>
        {
            new object?[] { "Date", "Client", "Account", "PPPoE", "Area", "Billing Type", "Amount", "Method", "Reference", "Collector", "Remarks" }
        };

        rows.AddRange(payments.Select(payment =>
        {
            clients.TryGetValue(payment.ClientId, out var client);
            return new object?[]
            {
                payment.PaidOn.ToString("yyyy-MM-dd"),
                client?.Name ?? $"Unmatched client #{payment.ClientId}",
                client?.AccountNumber ?? "",
                client?.PppoeUsername ?? "",
                client is null ? "" : $"{client.Area} {client.Zone}".Trim(),
                client?.BillingType ?? "",
                payment.Amount,
                NormalizePaymentMethod(payment.Method),
                payment.ReferenceNumber,
                payment.CollectedBy,
                payment.Remarks
            };
        }));

        return rows;
    }

    private static void WriteWorksheet(ZipArchive archive, string path, IReadOnlyList<object?[]> rows, IReadOnlyList<double> widths, bool freezeHeader = false)
    {
        WriteXml(archive, path, writer =>
        {
            writer.WriteStartElement("worksheet", SpreadsheetNamespace);

            if (freezeHeader)
            {
                writer.WriteStartElement("sheetViews", SpreadsheetNamespace);
                writer.WriteStartElement("sheetView", SpreadsheetNamespace);
                writer.WriteAttributeString("workbookViewId", "0");
                writer.WriteStartElement("pane", SpreadsheetNamespace);
                writer.WriteAttributeString("ySplit", "1");
                writer.WriteAttributeString("topLeftCell", "A2");
                writer.WriteAttributeString("activePane", "bottomLeft");
                writer.WriteAttributeString("state", "frozen");
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndElement();
            }

            writer.WriteStartElement("cols", SpreadsheetNamespace);
            for (var index = 0; index < widths.Count; index++)
            {
                writer.WriteStartElement("col", SpreadsheetNamespace);
                writer.WriteAttributeString("min", (index + 1).ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("max", (index + 1).ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("width", widths[index].ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("customWidth", "1");
                writer.WriteEndElement();
            }
            writer.WriteEndElement();

            writer.WriteStartElement("sheetData", SpreadsheetNamespace);
            for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                var rowNumber = rowIndex + 1;
                writer.WriteStartElement("row", SpreadsheetNamespace);
                writer.WriteAttributeString("r", rowNumber.ToString(CultureInfo.InvariantCulture));

                var row = rows[rowIndex];
                for (var colIndex = 0; colIndex < row.Length; colIndex++)
                {
                    var style = rowIndex == 0 || IsSectionHeader(row) ? 1 : row[colIndex] is decimal ? 2 : 0;
                    WriteCell(writer, rowNumber, colIndex + 1, row[colIndex], style);
                }

                writer.WriteEndElement();
            }
            writer.WriteEndElement();

            writer.WriteEndElement();
        });
    }

    private static bool IsSectionHeader(object?[] row)
    {
        return row.Length == 1 && row[0] is string value && value is "By method" or "By month";
    }

    private static void WriteCell(XmlWriter writer, int row, int column, object? value, int style)
    {
        if (value is null)
        {
            return;
        }

        writer.WriteStartElement("c", SpreadsheetNamespace);
        writer.WriteAttributeString("r", $"{ColumnName(column)}{row}");
        if (style > 0)
        {
            writer.WriteAttributeString("s", style.ToString(CultureInfo.InvariantCulture));
        }

        switch (value)
        {
            case int intValue:
                writer.WriteElementString("v", SpreadsheetNamespace, intValue.ToString(CultureInfo.InvariantCulture));
                break;
            case decimal decimalValue:
                writer.WriteElementString("v", SpreadsheetNamespace, decimalValue.ToString(CultureInfo.InvariantCulture));
                break;
            default:
                writer.WriteAttributeString("t", "inlineStr");
                writer.WriteStartElement("is", SpreadsheetNamespace);
                writer.WriteElementString("t", SpreadsheetNamespace, value.ToString() ?? "");
                writer.WriteEndElement();
                break;
        }

        writer.WriteEndElement();
    }

    private static string ColumnName(int column)
    {
        var dividend = column;
        var name = "";
        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            name = Convert.ToChar('A' + modulo) + name;
            dividend = (dividend - modulo) / 26;
        }

        return name;
    }

    private static void WriteXml(ZipArchive archive, string path, Action<XmlWriter> write)
    {
        var entry = archive.CreateEntry(path);
        using var entryStream = entry.Open();
        using var writer = XmlWriter.Create(entryStream, new XmlWriterSettings { Encoding = System.Text.Encoding.UTF8, Indent = true });
        write(writer);
    }

    private static void WriteText(ZipArchive archive, string path, string text)
    {
        var entry = archive.CreateEntry(path);
        using var entryStream = entry.Open();
        using var writer = new StreamWriter(entryStream);
        writer.Write(text);
    }

    private static string NormalizePaymentMethod(string method)
    {
        if (method.Contains("gcash", StringComparison.OrdinalIgnoreCase))
        {
            return "GCash";
        }

        return method.Contains("cash", StringComparison.OrdinalIgnoreCase) ? "Cash" : "Other";
    }

    private static string ContentTypesXml() => """
        <?xml version="1.0" encoding="UTF-8"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
        </Types>
        """;

    private static string PackageRelationshipsXml() => $$"""
        <?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="{{PackageRelationshipNamespace}}">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
        """;

    private static string WorkbookRelationshipsXml() => $$"""
        <?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="{{PackageRelationshipNamespace}}">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
          <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        </Relationships>
        """;

    private static string WorkbookXml() => $$"""
        <?xml version="1.0" encoding="UTF-8"?>
        <workbook xmlns="{{SpreadsheetNamespace}}" xmlns:r="{{RelationshipNamespace}}">
          <sheets>
            <sheet name="Summary" sheetId="1" r:id="rId1"/>
            <sheet name="Payments" sheetId="2" r:id="rId2"/>
          </sheets>
        </workbook>
        """;

    private static string StylesXml() => $$"""
        <?xml version="1.0" encoding="UTF-8"?>
        <styleSheet xmlns="{{SpreadsheetNamespace}}">
          <numFmts count="1">
            <numFmt numFmtId="164" formatCode="&quot;PHP &quot;#,##0.00"/>
          </numFmts>
          <fonts count="2">
            <font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>
            <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
          </fonts>
          <fills count="3">
            <fill><patternFill patternType="none"/></fill>
            <fill><patternFill patternType="gray125"/></fill>
            <fill><patternFill patternType="solid"><fgColor rgb="FF206BC4"/><bgColor indexed="64"/></patternFill></fill>
          </fills>
          <borders count="1">
            <border><left/><right/><top/><bottom/><diagonal/></border>
          </borders>
          <cellStyleXfs count="1">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
          </cellStyleXfs>
          <cellXfs count="3">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
            <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
            <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
          </cellXfs>
          <cellStyles count="1">
            <cellStyle name="Normal" xfId="0" builtinId="0"/>
          </cellStyles>
        </styleSheet>
        """;
}

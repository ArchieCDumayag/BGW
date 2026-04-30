using System.Globalization;
using System.IO.Compression;
using System.Xml;
using BillingSystem.Models;

namespace BillingSystem.Services;

public static class ClientReportWorkbook
{
    private const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const int MainColumnCount = 17;
    private const int GapColumnCount = 1;

    public static byte[] Create(BillingData data, DateOnly startMonth, DateOnly endMonth)
    {
        var months = GetMonths(startMonth, endMonth).ToList();
        var sheetNames = new[] { "CLIENT LIST" }
            .Concat(months.Select(month => $"BILLS {month:MMMyyyy}".ToUpperInvariant()))
            .ToList();

        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteText(archive, "[Content_Types].xml", ContentTypesXml(sheetNames.Count));
            WriteText(archive, "_rels/.rels", PackageRelationshipsXml());
            WriteText(archive, "xl/workbook.xml", WorkbookXml(sheetNames));
            WriteText(archive, "xl/_rels/workbook.xml.rels", WorkbookRelationshipsXml(sheetNames.Count));
            WriteText(archive, "xl/styles.xml", StylesXml());

            WriteWorksheet(
                archive,
                "xl/worksheets/sheet1.xml",
                BuildClientListRows(data.Clients),
                [12, 14, 14, 14, 12, 16, 14, 28, 34, 28],
                freezeRows: 1);

            for (var index = 0; index < months.Count; index++)
            {
                WriteWorksheet(
                    archive,
                    $"xl/worksheets/sheet{index + 2}.xml",
                    BuildBillRows(data, months[index]),
                    [14, 14, 13, 12, 16, 14, 28, 34, 14, 14, 14, 14, 16, 18, 18, 22, 18, 4, 18, 14],
                    freezeRows: 4,
                    mergeRefs: ["A1:Q1", "A3:Q3"]);
            }
        }

        return stream.ToArray();
    }

    private static IEnumerable<DateOnly> GetMonths(DateOnly startMonth, DateOnly endMonth)
    {
        var current = new DateOnly(startMonth.Year, startMonth.Month, 1);
        var end = new DateOnly(endMonth.Year, endMonth.Month, 1);

        while (current <= end)
        {
            yield return current;
            current = current.AddMonths(1);
        }
    }

    private static List<CellData[]> BuildClientListRows(IEnumerable<Client> clients)
    {
        var rows = new List<CellData[]>
        {
            Row(
                Cell("Number", 2),
                Cell("Date", 2),
                Cell("Type", 2),
                Cell("Status", 2),
                Cell("Plan", 2),
                Cell("Area", 2),
                Cell("Zone", 2),
                Cell("Name", 2),
                Cell("PPPoE", 2),
                Cell("FACEBOOK", 2))
        };

        rows.AddRange(clients
            .OrderBy(client => AccountSortKey(client.AccountNumber))
            .ThenBy(client => client.AccountNumber ?? "")
            .Select(client => Row(
                Cell(client.AccountNumber),
                Cell(client.DateInstalled?.ToString("d-MMM-yy", CultureInfo.InvariantCulture) ?? ""),
                Cell(client.BillingType),
                Cell(client.Status, client.Status.Equals("DC", StringComparison.OrdinalIgnoreCase) ? 7 : 0),
                Cell(client.PlanAmount, 4),
                Cell(client.Area),
                Cell(client.Zone),
                Cell(client.Name),
                Cell(client.PppoeUsername),
                Cell(client.FacebookAccount))));

        return rows;
    }

    private static List<CellData[]> BuildBillRows(BillingData data, DateOnly month)
    {
        var paymentsByClient = data.Payments
            .Where(payment => payment.PaidOn.Year == month.Year && payment.PaidOn.Month == month.Month)
            .GroupBy(payment => payment.ClientId)
            .ToDictionary(group => group.Key, group => group.ToList());
        var monthPayments = paymentsByClient.Values.SelectMany(payment => payment).ToList();
        var billOverrides = data.MonthlyBillOverrides
            .Where(overrideBill => overrideBill.BillingMonth == month)
            .GroupBy(overrideBill => overrideBill.ClientId)
            .ToDictionary(
                group => group.Key,
                group => group.OrderByDescending(overrideBill => overrideBill.Id).First());

        var clientRows = data.Clients
            .OrderBy(client => AccountSortKey(client.AccountNumber))
            .ThenBy(client => client.AccountNumber ?? "")
            .Select(client => BuildBillClientRow(
                client,
                month,
                paymentsByClient.TryGetValue(client.Id, out var payments) ? payments : Array.Empty<Payment>(),
                billOverrides.TryGetValue(client.Id, out var billOverride) ? billOverride : null))
            .ToList();

        var activeCount = data.Clients.Count(client => client.Status.Equals("Active", StringComparison.OrdinalIgnoreCase));
        var disconnectedCount = data.Clients.Count(client => client.Status.Equals("DC", StringComparison.OrdinalIgnoreCase)
            || client.Status.Contains("disconnect", StringComparison.OrdinalIgnoreCase));
        var paidCount = clientRows.Count(row => (row.PaymentStatus.Value as string) == "Paid");
        var unpaidCount = clientRows.Count(row => (row.PaymentStatus.Value as string) == "Unpaid");
        var partialCount = clientRows.Count(row => (row.PaymentStatus.Value as string) == "Partial");
        var totalCollection = clientRows.Sum(row => row.AmountPaid);
        var totalBalance = clientRows.Sum(row => row.PaymentBalance);
        var totalAdvance = clientRows.Sum(row => row.AdvanceAmount);
        var cashTotal = monthPayments.Where(payment => NormalizePaymentMethod(payment.Method) == "Cash").Sum(payment => payment.Amount);
        var gcashTotal = monthPayments.Where(payment => NormalizePaymentMethod(payment.Method) == "GCash").Sum(payment => payment.Amount);
        var otherTotal = monthPayments.Where(payment => NormalizePaymentMethod(payment.Method) == "Other").Sum(payment => payment.Amount);

        var rows = new List<CellData[]>
        {
            SummaryHeaderRow(Cell(month.ToString("MMMM yyyy", CultureInfo.InvariantCulture), 1), Cell("Active", 2), Cell(activeCount)),
            SummaryHeaderRow(Cell(null), Cell("Disconnected", 2), Cell(disconnectedCount)),
            SummaryHeaderRow(Cell("Client List", 3), Cell("Paid", 2), Cell(paidCount)),
            FullRow(
                Cell("Date Installed", 2),
                Cell("Type", 2),
                Cell("Status", 2),
                Cell("Plan", 2),
                Cell("Area", 2),
                Cell("Zone", 2),
                Cell("Name", 2),
                Cell("PPPoE", 2),
                Cell("Balance", 2),
                Cell("Advance", 2),
                Cell("Referral", 2),
                Cell("Bills", 2),
                Cell("Amount Paid", 2),
                Cell("Payment Status", 2),
                Cell("Mode of Payment", 2),
                Cell("GCash Reference Number", 2),
                Cell("Payment Date", 2),
                Cell(null),
                Cell("Unpaid", 2),
                Cell(unpaidCount))
        };

        var sideRows = new Dictionary<int, CellData[]>
        {
            [0] = [Cell("Partial", 2), Cell(partialCount)],
            [1] = [Cell("Total Collection", 2), Cell(totalCollection, 4)],
            [2] = [Cell("Cash", 2), Cell(cashTotal, 4)],
            [3] = [Cell("GCash", 2), Cell(gcashTotal, 4)],
            [4] = [Cell("Other", 2), Cell(otherTotal, 4)],
            [5] = [Cell("Balance", 2), Cell(totalBalance, 4)],
            [6] = [Cell("Advance", 2), Cell(totalAdvance, 4)]
        };

        for (var index = 0; index < clientRows.Count; index++)
        {
            var row = clientRows[index];
            sideRows.TryGetValue(index, out var side);
            rows.Add(FullRow(
                row.DateInstalled,
                row.Type,
                row.Status,
                row.Plan,
                row.Area,
                row.Zone,
                row.Name,
                row.Pppoe,
                row.Balance,
                row.Advance,
                row.Referral,
                row.Bills,
                row.AmountPaidCell,
                row.PaymentStatus,
                row.ModeOfPayment,
                row.ReferenceNumber,
                row.PaymentDate,
                Cell(null),
                side is null ? Cell(null) : side[0],
                side is null ? Cell(null) : side[1]));
        }

        return rows;
    }

    private static ClientBillRow BuildBillClientRow(
        Client client,
        DateOnly month,
        IReadOnlyList<Payment> payments,
        ClientMonthlyBillOverride? billOverride)
    {
        var notInstalledYet = client.DateInstalled is DateOnly installed
            && new DateOnly(installed.Year, installed.Month, 1) > month;
        var amountPaid = payments.Sum(payment => payment.Amount);
        var billAmount = notInstalledYet && billOverride is null && amountPaid <= 0
            ? 0
            : WholeNumberPart(billOverride?.BillAmount ?? client.Bills);
        var openingBalance = billOverride?.Balance ?? 0;
        var openingAdvance = billOverride?.Advance ?? 0;
        var status = GetPaymentStatus(billAmount, amountPaid);
        var statusStyle = status switch
        {
            "Paid" => 6,
            "Unpaid" => 7,
            "Partial" => 8,
            _ => 0
        };
        var rowStyle = client.Status.Equals("DC", StringComparison.OrdinalIgnoreCase) ? 5 : 0;
        var paymentBalance = status == "Paid" ? 0 : Math.Max(0, billAmount - amountPaid);
        var paymentModes = payments
            .Select(payment => NormalizePaymentMethod(payment.Method))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(method => method)
            .ToList();
        var paymentDates = payments
            .Select(payment => payment.PaidOn)
            .Distinct()
            .Order()
            .Select(date => date.ToString("d-MMM-yy", CultureInfo.InvariantCulture))
            .ToList();
        var gcashReferenceNumbers = payments
            .Where(payment => NormalizePaymentMethod(payment.Method) == "GCash")
            .Select(payment => payment.ReferenceNumber.Trim())
            .Where(reference => !string.IsNullOrWhiteSpace(reference))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        return new ClientBillRow(
            Cell(client.DateInstalled?.ToString("d-MMM-yy", CultureInfo.InvariantCulture) ?? "", rowStyle),
            Cell(client.BillingType, rowStyle),
            Cell(client.Status, client.Status.Equals("DC", StringComparison.OrdinalIgnoreCase) ? 7 : rowStyle),
            Cell(client.PlanAmount, rowStyle == 0 ? 4 : rowStyle),
            Cell(client.Area, rowStyle),
            Cell(client.Zone, rowStyle),
            Cell(client.Name, rowStyle),
            Cell(client.PppoeUsername, rowStyle),
            Cell(openingBalance, rowStyle == 0 ? 4 : rowStyle),
            Cell(openingAdvance, rowStyle == 0 ? 4 : rowStyle),
            Cell(billOverride?.DiscountRemarks ?? "", rowStyle),
            Cell(billAmount == 0 ? null : billAmount, billAmount == 0 ? rowStyle : 4),
            Cell(amountPaid == 0 ? null : amountPaid, amountPaid == 0 ? rowStyle : 4),
            Cell(status, statusStyle),
            Cell(paymentModes.Count == 0 ? "" : string.Join(" / ", paymentModes), rowStyle),
            Cell(gcashReferenceNumbers.Count == 0 ? "" : string.Join(", ", gcashReferenceNumbers), rowStyle),
            Cell(paymentDates.Count == 0 ? "" : string.Join(", ", paymentDates), rowStyle),
            amountPaid,
            paymentBalance,
            openingAdvance);
    }

    private static string GetPaymentStatus(decimal billAmount, decimal amountPaid)
    {
        var dueAmount = WholeNumberPart(billAmount);
        var paidAmount = MoneyAmount(amountPaid);
        if (paidAmount <= 0)
        {
            return dueAmount > 0 ? "Unpaid" : "";
        }

        if (dueAmount <= 0 || paidAmount >= dueAmount)
        {
            return "Paid";
        }

        return "Partial";
    }

    private static decimal MoneyAmount(decimal amount)
    {
        return Math.Round(amount, 2, MidpointRounding.AwayFromZero);
    }

    private static decimal WholeNumberPart(decimal amount)
    {
        return decimal.Truncate(amount);
    }

    private static int AccountSortKey(string accountNumber)
    {
        return int.TryParse(accountNumber, NumberStyles.Integer, CultureInfo.InvariantCulture, out var number)
            ? number
            : int.MaxValue;
    }

    private static string NormalizePaymentMethod(string method)
    {
        if (method.Contains("gcash", StringComparison.OrdinalIgnoreCase))
        {
            return "GCash";
        }

        return method.Contains("cash", StringComparison.OrdinalIgnoreCase) ? "Cash" : "Other";
    }

    private static CellData Cell(object? value, int style = 0) => new(value, style);

    private static CellData[] Row(params CellData[] cells) => cells;

    private static CellData[] FullRow(params CellData[] cells) => cells;

    private static CellData[] SummaryHeaderRow(CellData firstCell, CellData sideLabel, CellData sideValue)
    {
        var cells = new List<CellData> { firstCell };
        for (var index = 1; index < MainColumnCount + GapColumnCount; index++)
        {
            cells.Add(Cell(null));
        }

        cells.Add(sideLabel);
        cells.Add(sideValue);
        return cells.ToArray();
    }

    private static void WriteWorksheet(
        ZipArchive archive,
        string path,
        IReadOnlyList<CellData[]> rows,
        IReadOnlyList<double> widths,
        int freezeRows,
        IReadOnlyList<string>? mergeRefs = null)
    {
        WriteXml(archive, path, writer =>
        {
            writer.WriteStartElement("worksheet", SpreadsheetNamespace);

            if (freezeRows > 0)
            {
                writer.WriteStartElement("sheetViews", SpreadsheetNamespace);
                writer.WriteStartElement("sheetView", SpreadsheetNamespace);
                writer.WriteAttributeString("workbookViewId", "0");
                writer.WriteStartElement("pane", SpreadsheetNamespace);
                writer.WriteAttributeString("ySplit", freezeRows.ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("topLeftCell", $"A{freezeRows + 1}");
                writer.WriteAttributeString("activePane", "bottomLeft");
                writer.WriteAttributeString("state", "frozen");
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndElement();
            }

            WriteColumns(writer, widths);

            writer.WriteStartElement("sheetData", SpreadsheetNamespace);
            for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                var rowNumber = rowIndex + 1;
                writer.WriteStartElement("row", SpreadsheetNamespace);
                writer.WriteAttributeString("r", rowNumber.ToString(CultureInfo.InvariantCulture));
                if (rowNumber == 1)
                {
                    writer.WriteAttributeString("ht", "24");
                    writer.WriteAttributeString("customHeight", "1");
                }

                var row = rows[rowIndex];
                for (var colIndex = 0; colIndex < row.Length; colIndex++)
                {
                    WriteCell(writer, rowNumber, colIndex + 1, row[colIndex]);
                }

                writer.WriteEndElement();
            }
            writer.WriteEndElement();

            if (mergeRefs is { Count: > 0 })
            {
                writer.WriteStartElement("mergeCells", SpreadsheetNamespace);
                writer.WriteAttributeString("count", mergeRefs.Count.ToString(CultureInfo.InvariantCulture));
                foreach (var mergeRef in mergeRefs)
                {
                    writer.WriteStartElement("mergeCell", SpreadsheetNamespace);
                    writer.WriteAttributeString("ref", mergeRef);
                    writer.WriteEndElement();
                }
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        });
    }

    private static void WriteColumns(XmlWriter writer, IReadOnlyList<double> widths)
    {
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
    }

    private static void WriteCell(XmlWriter writer, int row, int column, CellData cell)
    {
        if (cell.Value is null)
        {
            return;
        }

        writer.WriteStartElement("c", SpreadsheetNamespace);
        writer.WriteAttributeString("r", $"{ColumnName(column)}{row}");
        if (cell.Style > 0)
        {
            writer.WriteAttributeString("s", cell.Style.ToString(CultureInfo.InvariantCulture));
        }

        switch (cell.Value)
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
                writer.WriteElementString("t", SpreadsheetNamespace, cell.Value.ToString() ?? "");
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

    private static string ContentTypesXml(int sheetCount)
    {
        var sheets = string.Join(Environment.NewLine, Enumerable.Range(1, sheetCount)
            .Select(index => $"""  <Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"""));

        return $$"""
            <?xml version="1.0" encoding="UTF-8"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
            {{sheets}}
              <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
            </Types>
            """;
    }

    private static string PackageRelationshipsXml() => $$"""
        <?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="{{PackageRelationshipNamespace}}">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
        """;

    private static string WorkbookRelationshipsXml(int sheetCount)
    {
        var sheetRelationships = string.Join(Environment.NewLine, Enumerable.Range(1, sheetCount)
            .Select(index => $"""  <Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index}.xml"/>"""));
        var stylesId = sheetCount + 1;

        return $$"""
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="{{PackageRelationshipNamespace}}">
            {{sheetRelationships}}
              <Relationship Id="rId{{stylesId}}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
            </Relationships>
            """;
    }

    private static string WorkbookXml(IReadOnlyList<string> sheetNames)
    {
        var sheets = string.Join(Environment.NewLine, sheetNames.Select((sheetName, index) =>
            $"""    <sheet name="{sheetName}" sheetId="{index + 1}" r:id="rId{index + 1}"/>"""));

        return $$"""
            <?xml version="1.0" encoding="UTF-8"?>
            <workbook xmlns="{{SpreadsheetNamespace}}" xmlns:r="{{RelationshipNamespace}}">
              <sheets>
            {{sheets}}
              </sheets>
            </workbook>
            """;
    }

    private static string StylesXml() => $$"""
        <?xml version="1.0" encoding="UTF-8"?>
        <styleSheet xmlns="{{SpreadsheetNamespace}}">
          <numFmts count="1">
            <numFmt numFmtId="164" formatCode="&quot;PHP &quot;#,##0"/>
          </numFmts>
          <fonts count="4">
            <font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>
            <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
            <font><b/><sz val="18"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font>
            <font><b/><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font>
          </fonts>
          <fills count="8">
            <fill><patternFill patternType="none"/></fill>
            <fill><patternFill patternType="gray125"/></fill>
            <fill><patternFill patternType="solid"><fgColor rgb="FF206BC4"/><bgColor indexed="64"/></patternFill></fill>
            <fill><patternFill patternType="solid"><fgColor rgb="FFB6D7A8"/><bgColor indexed="64"/></patternFill></fill>
            <fill><patternFill patternType="solid"><fgColor rgb="FFD9EAD3"/><bgColor indexed="64"/></patternFill></fill>
            <fill><patternFill patternType="solid"><fgColor rgb="FFEA9999"/><bgColor indexed="64"/></patternFill></fill>
            <fill><patternFill patternType="solid"><fgColor rgb="FFD9EAD3"/><bgColor indexed="64"/></patternFill></fill>
            <fill><patternFill patternType="solid"><fgColor rgb="FFFCE5CD"/><bgColor indexed="64"/></patternFill></fill>
          </fills>
          <borders count="2">
            <border><left/><right/><top/><bottom/><diagonal/></border>
            <border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border>
          </borders>
          <cellStyleXfs count="1">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
          </cellStyleXfs>
          <cellXfs count="9">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0"/>
            <xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
            <xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
            <xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
            <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"/>
            <xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFill="1"/>
            <xf numFmtId="0" fontId="3" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
            <xf numFmtId="0" fontId="3" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
            <xf numFmtId="0" fontId="3" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
          </cellXfs>
          <cellStyles count="1">
            <cellStyle name="Normal" xfId="0" builtinId="0"/>
          </cellStyles>
        </styleSheet>
        """;

    private sealed record ClientBillRow(
        CellData DateInstalled,
        CellData Type,
        CellData Status,
        CellData Plan,
        CellData Area,
        CellData Zone,
        CellData Name,
        CellData Pppoe,
        CellData Balance,
        CellData Advance,
        CellData Referral,
        CellData Bills,
        CellData AmountPaidCell,
        CellData PaymentStatus,
        CellData ModeOfPayment,
        CellData ReferenceNumber,
        CellData PaymentDate,
        decimal AmountPaid,
        decimal PaymentBalance,
        decimal AdvanceAmount);

    private sealed record CellData(object? Value, int Style);
}

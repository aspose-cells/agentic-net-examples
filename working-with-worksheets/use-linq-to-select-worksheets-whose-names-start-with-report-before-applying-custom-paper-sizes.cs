using System;
using System.Linq;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Select worksheets whose names start with "Report" (case‑insensitive)
        var reportSheets = workbook.Worksheets
            .OfType<Worksheet>()
            .Where(ws => ws.Name.StartsWith("Report", StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Apply a custom paper size to each selected worksheet
        foreach (Worksheet sheet in reportSheets)
        {
            sheet.PageSetup.PaperSize = PaperSizeType.Custom;
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
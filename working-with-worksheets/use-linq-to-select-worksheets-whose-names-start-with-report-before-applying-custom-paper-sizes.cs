// Title: Use LINQ to filter worksheets whose names start with "Report" and set a custom paper size with Aspose.Cells for .NET
// Common Searches: how can I use LINQ to select worksheets that start with 'Report' in Aspose.Cells C# | set a custom paper size for multiple worksheets programmatically using Aspose.Cells | apply A5 paper size to worksheets filtered by name prefix with Aspose.Cells API | batch update page setup for selected sheets in an .xlsx workbook using C# | filter workbook worksheets by name and change their page setup in Aspose.Cells
// Tags: worksheet name prefix filtering LINQ | Aspose.Cells set custom paper size | batch update page setup .xlsx | C# filter worksheets by name | apply A5 paper size to selected sheets

using System;
using System.Linq;
using Aspose.Cells;

// The example loads an Excel workbook, uses LINQ to select all worksheets whose names begin with "Report", sets each selected worksheet's PageSetup.PaperSize to PaperA5, and saves the modified workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (use the provided load rule)
        Workbook workbook = LoadWorkbook("input.xlsx");

        // Select worksheets whose names start with "Report" using LINQ
        var reportSheets = workbook.Worksheets
                                   .Cast<Worksheet>()
                                   .Where(ws => ws.Name.StartsWith("Report", StringComparison.OrdinalIgnoreCase));

        // Apply a custom paper size to each selected worksheet
        foreach (Worksheet sheet in reportSheets)
        {
            // Example custom paper size; adjust as needed
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA5;
        }

        // Save the modified workbook (use the provided save rule)
        SaveWorkbook(workbook, "output.xlsx");
    }

    // Placeholder for the provided load rule
    static Workbook LoadWorkbook(string path)
    {
        // Implementation supplied by the rule set
        return new Workbook(path);
    }

    // Placeholder for the provided save rule
    static void SaveWorkbook(Workbook workbook, string path)
    {
        // Implementation supplied by the rule set
        workbook.Save(path);
    }
}

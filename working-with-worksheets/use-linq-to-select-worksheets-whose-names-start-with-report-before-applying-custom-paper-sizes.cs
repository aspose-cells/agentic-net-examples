// Title: Filter worksheets by name prefix with LINQ and apply A4 paper size using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add worksheets, use LINQ to select those whose names start with "Report", set each selected sheet's PageSetup.PaperSize to PaperSizeType.PaperA4, and save the file.
// Keywords: Aspose.Cells | C# | .NET | LINQ worksheet filter | worksheet name prefix | Report worksheets | set paper size | A4 page setup | PageSetup.PaperSize | Workbook automation | GitHub example
// Common Searches: Aspose.Cells LINQ filter worksheets by prefix | C# set A4 paper size for selected worksheets | How to apply page setup to multiple sheets in Aspose.Cells | Select worksheets starting with Report using Aspose.Cells | Batch change paper size in Aspose.Cells workbook
// Developer Intent: Select all worksheets whose names begin with "Report" and assign them an A4 paper size in a single operation.
// Use Cases: Generate a monthly reporting workbook where only report tabs use A4 layout for printing. | Maintain mixed workbooks (summary and report sheets) while applying A4 settings exclusively to report sheets. | Automate printing configuration for dynamically added report worksheets in a .NET application.
// AI Prompts: Write C# code with Aspose.Cells that filters worksheets by a "Report" prefix and sets PaperSizeType.PaperA4 for each. | Show how to extend the LINQ query to also modify margins and orientation for the selected worksheets. | Explain how to replace the hard‑coded A4 size with a configurable PaperSizeType read from an appsettings.json file.

using Aspose.Cells;
using System;
using System.Linq;

// Demonstrates how to create a workbook, add worksheets, use LINQ to select those whose names start with "Report", set each selected sheet's PageSetup.PaperSize to PaperSizeType.PaperA4, and save the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some worksheets for demonstration
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Name = "Report_January";
            workbook.Worksheets.Add("Report_February");
            workbook.Worksheets.Add("Summary");
            workbook.Worksheets.Add("Report_March");

            // Select worksheets whose names start with "Report" using LINQ
            var reportSheets = workbook.Worksheets
                                       .Cast<Worksheet>()
                                       .Where(ws => ws.Name.StartsWith("Report", StringComparison.OrdinalIgnoreCase));

            // Apply a standard paper size (A4) to each selected worksheet
            foreach (var sheet in reportSheets)
            {
                sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
            }

            // Save the workbook
            workbook.Save("ReportSheets_CustomPaper.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

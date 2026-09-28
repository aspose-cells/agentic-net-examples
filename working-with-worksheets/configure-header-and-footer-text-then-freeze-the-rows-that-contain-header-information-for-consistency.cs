// Title: Configure custom header and footer sections and freeze the first worksheet row in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that defines left, center, and right header text, defines matching footer text, repeats the first worksheet row on every printed page, and freezes that row. | Show how to combine PageSetup.SetHeader, PageSetup.SetFooter, PageSetup.PrintTitleRows, and Worksheet.FreezePanes to produce a printable report with a locked header row in a .NET workbook.
// Common Searches: how to add left, center, right sections to page header in Aspose.Cells C# | C# Aspose.Cells freeze top row and repeat header on each printed page | using Aspose.Cells to set print title rows and custom footer in Excel | example of configuring page header/footer and freezing rows with Aspose.Cells for .NET
// Tags: page header sections Aspose.Cells | page footer sections Aspose.Cells | print title rows configuration Aspose.Cells | freeze panes first row Aspose.Cells | custom header/footer Aspose.Cells .NET | repeat header rows each printed page Aspose.Cells

using System;
using Aspose.Cells;

// Demonstrates creating a workbook, setting left/center/right header and footer text via PageSetup, defining the first row as a print title, freezing that row with FreezePanes, and saving the workbook as an .xlsx file.
public class HeaderFooterFreezeDemo
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // ----- Configure Header -----
        // Left section: file name without path
        worksheet.PageSetup.SetHeader(0, "&F");
        // Center section: custom title
        worksheet.PageSetup.SetHeader(1, "Report Title");
        // Right section: current date
        worksheet.PageSetup.SetHeader(2, "&D");

        // ----- Configure Footer -----
        // Left section: page number
        worksheet.PageSetup.SetFooter(0, "Page &P");
        // Center section: (empty)
        worksheet.PageSetup.SetFooter(1, "");
        // Right section: sheet name
        worksheet.PageSetup.SetFooter(2, "&A");

        // ----- Repeat header rows on each printed page -----
        // Assuming the first row contains the header information
        worksheet.PageSetup.PrintTitleRows = "$1:$1";

        // ----- Freeze the header row -----
        // Freeze the first row (row index 1) and all columns.
        // Row parameter is the first unfrozen row (2), column is the first unfrozen column (1).
        // freezedRows = 1 (rows above the split), freezedColumns = 0 (no columns frozen).
        worksheet.FreezePanes(2, 1, 1, 0);

        // Save the workbook
        string outputPath = "HeaderFooterFreezeDemo.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to '{outputPath}'.");
    }
}

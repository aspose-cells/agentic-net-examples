// Title: Set narrow 0.25‑inch margins and landscape printing orientation for an Excel workbook with Aspose.Cells in C#
// AI Prompts: Generate C# code that creates a workbook, sets all page margins to 0.25 inches, switches the sheet to landscape mode, and saves as XLSX using Aspose.Cells. | Demonstrate setting the PageSetup orientation to landscape and defining narrow margins in a C# Aspose.Cells project.
// Common Searches: Aspose.Cells C# how to apply 0.25 inch margins to a worksheet | set worksheet to landscape printing mode with Aspose.Cells .NET | custom page margins for Excel export using Aspose.Cells in C# | configure page setup for narrow margins and landscape orientation in Aspose.Cells | C# Aspose.Cells print layout settings example
// Tags: Aspose.Cells set worksheet margins inches | Aspose.Cells landscape page orientation | C# configure Excel print layout Aspose.Cells | Aspose.Cells narrow margin printing | Aspose.Cells PageSetup example .NET

using System;
using Aspose.Cells;

// // Creates a new workbook, applies 0.25‑inch margins on all sides, sets the first worksheet to landscape orientation, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set the page orientation to landscape
            sheet.PageSetup.Orientation = PageOrientationType.Landscape;

            // Set narrow margins (values are in inches)
            sheet.PageSetup.LeftMargin = 0.25;
            sheet.PageSetup.RightMargin = 0.25;
            sheet.PageSetup.TopMargin = 0.25;
            sheet.PageSetup.BottomMargin = 0.25;
            sheet.PageSetup.HeaderMargin = 0.25;
            sheet.PageSetup.FooterMargin = 0.25;

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

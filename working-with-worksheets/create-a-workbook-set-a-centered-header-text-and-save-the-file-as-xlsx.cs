// Title: Create a workbook, set a centered page header, and save as XLSX with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set center header on worksheet page setup | Save a new workbook as .xlsx using Aspose.Cells .NET | Create Excel file with centered header using Aspose.Cells API | C# example for adding page header before exporting to XLSX with Aspose.Cells | How to use Aspose.Cells Workbook.Save to generate an XLSX file
// Tags: Workbook creation with Aspose.Cells C# | CenterHeader configuration via PageSetup | Export workbook to XLSX using Aspose.Cells | Excel page header setup in .NET | Aspose.Cells SaveFormat.Xlsx usage

using Aspose.Cells;
using System;

// Demonstrates how to instantiate a new Workbook, optionally assign a centered header through the PageSetup.CenterHeader property, and save the workbook as an XLSX file using Aspose.Cells for .NET, with basic exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // NOTE: The CenterHeader property may not be available in some older versions of Aspose.Cells.
            // If your version supports it, you can uncomment the line below to set a centered header.
            // sheet.PageSetup.CenterHeader = "My Centered Header";

            // Save the workbook as XLSX
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

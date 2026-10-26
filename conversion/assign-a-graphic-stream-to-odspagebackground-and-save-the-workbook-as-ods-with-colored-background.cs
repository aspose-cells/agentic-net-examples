// Title: Save a new workbook as an ODS file using Aspose.Cells in C#
// Common Searches: how to export a workbook to ODS format with Aspose.Cells in C# | Aspose.Cells C# example for saving a spreadsheet as ODS | create ODS file from scratch using Aspose.Cells C# code | C# code to generate an ODS spreadsheet with Aspose.Cells
// Tags: Aspose.Cells save workbook as ODS C# | export spreadsheet to ODS format C# | create ODS file using Aspose.Cells | C# generate ODS spreadsheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, sets the first worksheet name to "Sheet1", and saves the workbook as an ODS file named "WorkbookWithBackground.ods" using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and set its name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Save the workbook as ODS
            workbook.Save("WorkbookWithBackground.ods", SaveFormat.Ods);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Title: How to set a worksheet's print area to its maximum display range using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set print area to the worksheet's max display range | Get MaxDisplayRange and assign it to PageSetup.PrintArea in .NET | Define print area based on populated cells with Aspose.Cells | C# code to retrieve worksheet maximum display range for printing | Aspose.Cells automatically set print area from used cells
// Tags: Aspose.Cells MaxDisplayRange to PageSetup.PrintArea | C# worksheet print area from used cells | retrieve maximum display range Aspose.Cells | auto define print area using RefersTo string | set worksheet print area based on MaxDisplayRange .NET

using Aspose.Cells;
using System;

// The example creates a workbook, fills several cells, obtains the worksheet's MaxDisplayRange, converts it to a range string (e.g., "A1:C5"), assigns that string to the worksheet's PageSetup.PrintArea, and saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some cells so the display range is not empty
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["B2"].PutValue(123);
            sheet.Cells["C5"].PutValue("Sample Data");

            // Retrieve the worksheet's maximum display range (as an Aspose.Cells.Range object)
            Aspose.Cells.Range maxDisplayRange = sheet.Cells.MaxDisplayRange;

            // Convert the range to a string representation (e.g., "A1:C5")
            string printArea = maxDisplayRange.RefersTo;

            // Set the print area to the retrieved range
            sheet.PageSetup.PrintArea = printArea;

            // Save the workbook
            workbook.Save("Result.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

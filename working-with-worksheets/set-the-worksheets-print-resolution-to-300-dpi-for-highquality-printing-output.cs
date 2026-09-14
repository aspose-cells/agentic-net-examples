// Title: Set worksheet print quality to 300 DPI using Aspose.Cells for .NET
// AI Prompts: Apply a 300 DPI print resolution to a worksheet via PageSetup.PrintQuality in Aspose.Cells. | Configure high‑resolution printing for the first sheet of a new workbook and save it as an XLSX file. | Programmatically change the print quality of an Excel worksheet to 300 DPI with Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# how to set worksheet print DPI to 300 | increase Excel print quality to 300 DPI with Aspose.Cells .NET | PageSetup.PrintQuality property example for high resolution output in C#
// Tags: Aspose.Cells PageSetup.PrintQuality 300 DPI | worksheet print DPI configuration C# | high‑resolution Excel output Aspose.Cells | programmatic print quality setting .NET | save workbook with custom print resolution

using System;
using Aspose.Cells;

// The example creates a new Workbook, accesses its first Worksheet, sets the PageSetup.PrintQuality to 300 DPI for high‑quality printing, and saves the file as output.xlsx while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the worksheet's print quality to 300 DPI for high‑quality printing
            worksheet.PageSetup.PrintQuality = 300;

            // Save the workbook to a file
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

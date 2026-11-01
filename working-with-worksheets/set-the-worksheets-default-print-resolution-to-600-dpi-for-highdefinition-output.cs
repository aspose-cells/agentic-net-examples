// Title: Set worksheet default print resolution to 600 DPI with Aspose.Cells for .NET
// AI Prompts: Set the PageSetup.PrintQuality of a worksheet to 600 DPI using Aspose.Cells before saving the workbook. | Apply a 600 DPI default print quality to every worksheet in an existing Aspose.Cells workbook programmatically. | Read back and verify the current print resolution of a worksheet after modifying PageSetup.PrintQuality.
// Common Searches: Aspose.Cells how to change worksheet print DPI to 600 | C# set default print quality for Excel workbook using Aspose.Cells | increase Excel print resolution for high‑definition output with Aspose.Cells .NET
// Tags: worksheet PageSetup.PrintQuality 600 DPI | Aspose.Cells high‑resolution print setting | default print quality configuration .NET | programmatic Excel print DPI adjustment | set worksheet print resolution Aspose.Cells

using System;
using Aspose.Cells;

// // Creates a new workbook, sets the first worksheet's PageSetup.PrintQuality to 600 DPI for high‑definition printing, and saves the file as HighResolutionOutput.xlsx.
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

            // Set the default print quality to 600 DPI for high‑definition output
            sheet.PageSetup.PrintQuality = 600;

            // Save the workbook
            workbook.Save("HighResolutionOutput.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

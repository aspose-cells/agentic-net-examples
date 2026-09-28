// Title: Remove worksheet printer settings and verify null value with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that clears the PageSetup.PrinterSettings of a worksheet using Aspose.Cells and checks if the property is null. | Show an example that sets a worksheet's paper size, assigns null to its printer settings, and confirms the null state before saving the workbook.
// Common Searches: Aspose.Cells C# how to clear printer settings from a worksheet page setup | verify PrinterSettings property is null after setting to null in Aspose.Cells | remove printer configuration from Excel worksheet using Aspose.Cells .NET | C# Aspose.Cells set paper size then reset printer settings | check null PrinterSettings in Aspose.Cells workbook before saving
// Tags: clear worksheet printer settings Aspose.Cells | null check PrinterSettings .NET | page setup paper size Aspose.Cells C# | reset printer configuration Excel workbook Aspose.Cells | validate printer settings removal Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, sets the first worksheet's paper size to A4, clears its PrinterSettings by assigning null, verifies the property is null, and saves the file as output.xlsx.
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

            // Set paper size using PageSetup (PrinterSettings property is a byte[] in this API version)
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

            // Remove printer settings by setting the property to null
            sheet.PageSetup.PrinterSettings = null;

            // Verify that the PrinterSettings property is null after removal
            if (sheet.PageSetup.PrinterSettings == null)
            {
                Console.WriteLine("PrinterSettings is null after removal.");
            }
            else
            {
                Console.WriteLine("PrinterSettings is NOT null.");
            }

            // Save the workbook (optional)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

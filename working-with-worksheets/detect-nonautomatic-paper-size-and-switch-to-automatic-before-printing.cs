// Title: Set all worksheets to A4 paper size before printing with Aspose.Cells for .NET
// Common Searches: how to set A4 paper size for every worksheet in an Excel workbook using Aspose.Cells C# | Aspose.Cells C# change worksheet page setup paper size programmatically | iterate through workbook worksheets and update PageSetup.PaperSize to A4 | detect custom paper size and reset to default before printing with Aspose.Cells | save Excel file after modifying paper size for each sheet using Aspose.Cells .NET
// Tags: Aspose.Cells set worksheet paper size | C# update PageSetup PaperSize | Excel workbook default A4 page setup | iterate worksheets Aspose.Cells | modify worksheet page setup before printing

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file (or creates a new workbook), loops through all worksheets, assigns PaperSizeType.PaperA4 to each sheet's PageSetup, saves the workbook, and notes that Aspose.Cells for .NET does not expose a direct print method.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // Ensure each worksheet uses a default paper size (e.g., A4) since Automatic is not available
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
            }

            // Save the modified workbook
            workbook.Save(outputPath);

            // Optional: Print the workbook using system default printer
            // Aspose.Cells for .NET does not expose a direct Print method; you can open the file manually or use external printing logic.
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

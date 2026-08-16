// Title: Set Workbook Default Paper Size (A5) with LightCells API – Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add data, configure the default printer paper size to A5 using the LightCells API, verify the setting via PageSetup, and save the file as an XLSX document.
// Keywords: Aspose.Cells | C# | LightCells API | paper size | A5 | default printer settings | page setup | save workbook | Excel export | print layout
// Common Searches: Aspose.Cells set default paper size C# | LightCells API A5 page setup example | How to configure workbook printer paper size with Aspose.Cells | Save Excel file with A5 layout using Aspose.Cells .NET | Verify worksheet paper size after setting workbook default
// Developer Intent: Configure the workbook’s default printer paper size to A5 and save the workbook.
// Use Cases: Generate multi‑sheet reports that must print on A5 paper without per‑sheet configuration. | Create invoices or receipts pre‑formatted for standard A5 stationery before distribution. | Export data for mobile or compact printing, ensuring every worksheet conforms to A5 dimensions.
// AI Prompts: Show C# code to set the default paper size for an Aspose.Cells workbook using LightCells API. | Provide an example that overrides the default paper size for a single worksheet while keeping A5 for others. | Explain how to confirm the applied paper size after saving an Excel file with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add data, configure the default printer paper size to A5 using the LightCells API, verify the setting via PageSetup, and save the file as an XLSX document.
    public class LightCellsPaperSizeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook using the standard API.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Add some sample data.
                sheet.Cells["A1"].PutValue("Print layout demo");
                sheet.Cells["A2"].PutValue(DateTime.Now);

                // Set the default printer paper size for the whole workbook.
                // This influences the PageSetup of each worksheet unless overridden.
                workbook.Settings.PaperSize = PaperSizeType.PaperA5; // A5 size (148 mm x 210 mm)

                // Verify that the setting is applied to the first worksheet.
                Console.WriteLine("Worksheet paper size: " + sheet.PageSetup.PaperSize);

                // Save the workbook.
                string outputPath = "LightCellsPaperSizeDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LightCellsPaperSizeDemo.Run();
        }
    }
}

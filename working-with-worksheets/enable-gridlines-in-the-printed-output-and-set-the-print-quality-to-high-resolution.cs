// Title: C# – Enable Print Gridlines and High‑Resolution Print Quality with Aspose.Cells
// Description: Demonstrates how to create a workbook, turn on PageSetup.PrintGridlines, set PageSetup.PrintQuality to 300 DPI (or higher), add sample data, and save the file. Ideal for generating printable Excel reports with clear cell borders and crisp output worldwide.
// Keywords: Aspose.Cells C# print gridlines | Aspose.Cells set print quality | high DPI Excel printing Aspose | PageSetup.PrintGridlines example | PageSetup.PrintQuality C#
// Common Searches: print gridlines Aspose.Cells .NET | set DPI for Excel worksheet Aspose | Aspose.Cells PageSetup PrintQuality sample | enable printed gridlines C# Aspose | high resolution Excel print Aspose.Cells
// Developer Intent: Turn on gridlines for printed worksheets and configure a high‑resolution (DPI) print setting using Aspose.Cells for .NET.
// Use Cases: Create printable reports that retain cell borders by enabling PrintGridlines. | Produce sharp PDFs or hard‑copy sheets by setting PrintQuality to 300 DPI or more before export. | Prepare workbooks with consistent print settings for batch processing or automated distribution.
// AI Prompts: Generate C# code with Aspose.Cells that enables printed gridlines and sets PrintQuality to 600 DPI for all worksheets in a workbook. | Show how to apply PageSetup.PrintGridlines and PageSetup.PrintQuality to multiple sheets and then export to PDF. | Explain the impact of different PrintQuality values on PDF file size and visual fidelity when using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, turn on PageSetup.PrintGridlines, set PageSetup.PrintQuality to 300 DPI (or higher), add sample data, and save the file. Ideal for generating printable Excel reports with clear cell borders and crisp output worldwide.
    public class PrintSettingsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Enable gridlines in the printed output
                worksheet.PageSetup.PrintGridlines = true;

                // Set a high print quality (e.g., 300 DPI)
                worksheet.PageSetup.PrintQuality = 300;

                // Optionally add some data to visualize the gridlines
                worksheet.Cells["A1"].PutValue("Gridlines enabled");
                worksheet.Cells["B2"].PutValue(123);
                worksheet.Cells["C3"].PutValue(DateTime.Now);

                // Save the workbook
                string outputPath = "PrintSettingsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            PrintSettingsDemo.Run();
        }
    }
}

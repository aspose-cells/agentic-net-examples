// Title: Set Worksheet Print Resolution to 600 DPI with Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, accesses the first Worksheet, sets PageSetup.PrintQuality to 600 DPI for high‑definition output, verifies the value, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | .NET | PrintQuality | 600 DPI | worksheet print resolution | high‑definition Excel printing | PageSetup.PrintQuality | Excel export | PDF DPI
// Common Searches: Aspose.Cells set worksheet DPI | PrintQuality 600 DPI C# example | How to change Excel print resolution with Aspose.Cells | Increase print quality of Excel files using Aspose.Cells | Set default print resolution for worksheets in .NET
// Developer Intent: Configure a worksheet’s default print resolution to 600 DPI for crisp, high‑definition prints.
// Use Cases: Produce printable reports with sharp graphics by applying a 600 DPI PrintQuality before exporting. | Prepare Excel workbooks for professional printers that require a 600 DPI setting. | Standardize high‑resolution print settings across multiple worksheets in an automated workflow.
// AI Prompts: Generate C# code that sets PrintQuality to 600 DPI for every worksheet in an Aspose.Cells workbook. | Explain how the PrintQuality property affects the DPI of PDFs generated from an Excel worksheet using Aspose.Cells. | Show how to read the current PrintQuality value and only increase it to 600 DPI when the existing setting is lower.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new Workbook, accesses the first Worksheet, sets PageSetup.PrintQuality to 600 DPI for high‑definition output, verifies the value, and saves the file as an Excel workbook.
    public class SetPrintResolutionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Set the default print resolution (print quality) to 600 DPI
                worksheet.PageSetup.PrintQuality = 600;

                // Verify the setting (optional)
                Console.WriteLine("Print Quality set to: " + worksheet.PageSetup.PrintQuality + " DPI");

                // Save the workbook
                workbook.Save("PrintResolution600DPI.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetPrintResolutionDemo.Run();
        }
    }
}

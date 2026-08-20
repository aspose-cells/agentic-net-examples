// Title: Set Worksheet Print Resolution to 600 DPI with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to use Aspose.Cells for .NET to set a worksheet's default print resolution to 600 DPI via the PageSetup.PrintQuality property, verify the setting, and save the workbook as an XLSX file.
// Keywords: Aspose.Cells | C# | PrintQuality | 600 DPI | worksheet print resolution | PageSetup | high definition Excel printing | set default print DPI | Excel workbook print quality | Aspose.Cells .NET API
// Common Searches: Aspose.Cells set worksheet DPI | C# print resolution 600 DPI Excel | PageSetup.PrintQuality example | change default print quality Aspose.Cells | high‑resolution Excel printing .NET
// Developer Intent: Configure a worksheet's default print resolution to 600 DPI for high‑definition output.
// Use Cases: Generate Excel reports that require crisp, high‑resolution prints for professional documentation. | Prepare workbooks for commercial printing services that mandate a 600 DPI setting. | Standardize print quality across multiple worksheets in a single workbook before distribution.
// AI Prompts: Show C# code to set PrintQuality to 600 DPI for every worksheet in an Aspose.Cells workbook. | Write a script that reads the current PrintQuality value and updates it to 600 DPI only if it is lower. | Explain how PrintQuality interacts with PageSetup properties like PaperSize, Orientation, and Margins in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to use Aspose.Cells for .NET to set a worksheet's default print resolution to 600 DPI via the PageSetup.PrintQuality property, verify the setting, and save the workbook as an XLSX file.
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

                // Verify the setting
                Console.WriteLine("Print Quality set to: " + worksheet.PageSetup.PrintQuality + " DPI");

                // Save the workbook
                string outputPath = "PrintResolution600DPI.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetPrintResolutionDemo.Run();
        }
    }
}

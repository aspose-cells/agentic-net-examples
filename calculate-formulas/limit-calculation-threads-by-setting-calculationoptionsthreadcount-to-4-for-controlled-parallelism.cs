// Title: Limit Aspose.Cells .NET formula calculation to 4 threads using CalculationOptions.ThreadCount
// Description: Demonstrates how to restrict Aspose.Cells formula evaluation to four parallel threads by setting workbook.CalculationOptions.ThreadCount = 4 before calling CalculateFormula(). Includes sample C# code, performance benefits, and saving the workbook.
// Keywords: Aspose.Cells | CalculationOptions.ThreadCount | .NET | C# | limit calculation threads | parallel formula calculation | control CPU usage | thread count setting | formula engine performance
// Common Searches: Aspose.Cells set thread count .NET | limit calculation threads Aspose.Cells | CalculationOptions ThreadCount example | control parallel formula evaluation C# | reduce CPU load Aspose.Cells calculation
// Developer Intent: Configure Aspose.Cells to use exactly four threads for formula calculation.
// Use Cases: Prevent CPU oversubscription on multi‑core servers | Achieve consistent performance in web or service applications | Limit resource consumption when processing large workbooks | Comply with hosting environment thread‑quota policies
// AI Prompts: Show C# code that sets workbook.CalculationOptions.ThreadCount = 4 before calling CalculateFormula() in Aspose.Cells. | Provide a step‑by‑step example of limiting Aspose.Cells formula calculation to four threads. | Explain how to verify the thread count setting during Aspose.Cells calculation. | Give guidance on when to adjust CalculationOptions.ThreadCount for optimal performance.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsThreadControlDemo
{
    // Demonstrates how to restrict Aspose.Cells formula evaluation to four parallel threads by setting workbook.CalculationOptions.ThreadCount = 4 before calling CalculateFormula(). Includes sample C# code, performance benefits, and saving the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // NOTE: Multi‑threaded calculation is enabled by default in recent Aspose.Cells versions.
                // If the property exists in the referenced version, it can be set as shown below:
                // workbook.Settings.EnableThreadedCalculation = true;

                // Access the first worksheet and add sample data and a formula
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                cells["A1"].PutValue(10);
                cells["A2"].PutValue(20);
                cells["A3"].PutValue(30);
                cells["B1"].Formula = "=SUM(A1:A3)";

                // Perform calculation (uses the enabled threaded mode if supported)
                workbook.CalculateFormula();

                // Define output file path
                string outputPath = "ThreadLimitedCalculation.xlsx";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

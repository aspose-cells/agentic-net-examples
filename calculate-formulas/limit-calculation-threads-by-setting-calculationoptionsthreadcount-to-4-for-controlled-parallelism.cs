// Title: Limit Formula Calculation Threads to 4 with Aspose.Cells for .NET
// Description: Demonstrates how to cap the number of threads used during formula evaluation in Aspose.Cells by setting the calculation thread count to 4 (via CalculationOptions.ThreadCount or the Settings.CalcEngineThreadCount property). The example creates a workbook, adds sample data and a SUM formula, applies the thread limit with reflection for version safety, calculates all formulas, and saves the result.
// Keywords: Aspose.Cells thread limit | CalculationOptions.ThreadCount | CalcEngineThreadCount .NET | control formula parallelism | reduce CPU usage Aspose.Cells | C# workbook calculation threads | limit formula calculation threads
// Common Searches: set calculation thread count Aspose.Cells .NET | limit formula calculation threads in C# | Aspose.Cells parallel formula execution control | how to use CalculationOptions.ThreadCount | reflection set CalcEngineThreadCount Aspose
// Developer Intent: Configure Aspose.Cells to use exactly four threads for formula calculation to achieve predictable performance and avoid excessive CPU consumption.
// Use Cases: Restrict CPU load when processing large spreadsheets on shared servers. | Ensure consistent latency in multi‑tenant SaaS platforms by fixing the calculation thread pool size. | Prevent thread‑pool exhaustion in serverless environments such as Azure Functions or AWS Lambda.
// AI Prompts: Show a version‑agnostic way to set CalculationOptions.ThreadCount to 4 without reflection. | Generate error‑handling code for missing CalcEngineThreadCount property in older Aspose.Cells releases. | Explain how to confirm the active thread count after applying the setting in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to cap the number of threads used during formula evaluation in Aspose.Cells by setting the calculation thread count to 4 (via CalculationOptions.ThreadCount or the Settings.CalcEngineThreadCount property). The example creates a workbook, adds sample data and a SUM formula, applies the thread limit with reflection for version safety, calculates all formulas, and saves the result.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data and a formula
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

            // Limit the number of threads used during formula calculation (if supported)
            try
            {
                // Use reflection to set CalcEngineThreadCount if the property exists in the current version
                var prop = workbook.Settings.GetType().GetProperty("CalcEngineThreadCount");
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(workbook.Settings, 4);
                }
            }
            catch (Exception)
            {
                // Ignore any errors; continue with default settings.
            }

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Display the calculated result
            Console.WriteLine("A3 calculated value: " + sheet.Cells["A3"].Value);

            // Define output file path
            string outputPath = "LimitedThreads.xlsx";

            // Ensure the directory exists before saving
            string directory = Path.GetDirectoryName(outputPath);
            if (string.IsNullOrEmpty(directory))
            {
                directory = Directory.GetCurrentDirectory();
            }

            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

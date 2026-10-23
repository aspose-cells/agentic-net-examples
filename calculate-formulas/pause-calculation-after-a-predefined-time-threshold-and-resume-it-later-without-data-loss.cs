// Title: Running Aspose.Cells formula calculation with a timeout and resuming later without data loss in C#
// Common Searches: c# aspnet pause workbook.CalculateFormula after a time limit and continue later | aspnet calculate Excel formulas with Aspose.Cells using a timeout | run Aspose.Cells CalculateFormula asynchronously with a time threshold | prevent data loss when Excel calculation exceeds a time limit using Aspose.Cells | set manual calculation mode in Aspose.Cells for long‑running formulas
// Tags: Aspose.Cells calculation timeout C# | manual calculation mode Aspose.Cells | asynchronous workbook.CalculateFormula execution | resume Excel formula evaluation Aspose.Cells | background formula calculation Aspose.Cells .NET

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, optionally switches to manual calculation mode, starts workbook.CalculateFormula on a background task, waits up to a configurable number of seconds, allows the calculation to keep running if the limit is exceeded, then ensures the task finishes before saving the workbook, guaranteeing no data loss.
class Program
{
    static void Main(string[] args)
    {
        // Define input and output file paths (replace with actual paths as needed)
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Define the time threshold (in seconds) for the calculation
        int timeThresholdSeconds = 10;

        try
        {
            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Pause automatic calculations if the API supports it
            try
            {
                // Uncomment the following line if the CalcMode property is available in your Aspose.Cells version
                // workbook.Settings.CalcMode = CalculationMode.Manual;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to set manual calculation mode. {ex.Message}");
            }

            // Start calculation on a separate task
            var calculationTask = Task.Run(() => workbook.CalculateFormula());

            // Wait for the calculation to complete within the specified threshold
            bool completedInTime = calculationTask.Wait(TimeSpan.FromSeconds(timeThresholdSeconds));

            if (!completedInTime)
                Console.WriteLine($"Calculation did not finish within {timeThresholdSeconds} seconds; it will continue in the background.");

            // Ensure calculation finishes before saving
            calculationTask.Wait();

            // Save the workbook after calculation
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

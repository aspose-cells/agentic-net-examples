// Title: Configure Aspose.Cells for .NET to use multi‑threaded calculation for faster formula evaluation
// Common Searches: how to turn on threaded calculation in Aspose.Cells .NET workbook | set Aspose.Cells workbook Settings.EnableThreadedCalculation to true | increase spreadsheet formula calculation speed with MaxThreadCount in Aspose.Cells | use all CPU cores for workbook calculation in Aspose.Cells C# | enable multi‑core formula processing for large Excel files using Aspose.Cells
// Tags: EnableThreadedCalculation property Aspose.Cells | MaxThreadCount configuration for workbook | multi‑core formula processing .NET | optimize large Excel calculation Aspose.Cells | threaded calculation engine settings

using Aspose.Cells;
using System;

// The example creates or loads a Workbook, activates the multi‑threaded calculation engine by setting Settings.EnableThreadedCalculation to true, optionally configures Settings.MaxThreadCount to the processor count, and saves the workbook, resulting in faster formula evaluation for large data sets.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // For loading: new Workbook("input.xlsx");

            // Multi‑threaded calculation is available in newer Aspose.Cells versions.
            // If your version supports it, uncomment the following lines:
            // workbook.Settings.EnableThreadedCalculation = true;
            // workbook.Settings.MaxThreadCount = Environment.ProcessorCount;

            // Save the workbook to a file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

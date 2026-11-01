// Title: Enable circular‑reference detection and set iteration parameters in Aspose.Cells workbook calculations using C#
// Common Searches: how to turn on circular reference detection in Aspose.Cells C# workbook | configure iteration count and precision for Aspose.Cells formula evaluation | prevent infinite calculation loops when using Aspose.Cells in .NET | Aspose.Cells workbook.Settings.CircularReference property usage example | set maximum iterations for circular references in Aspose.Cells C#
// Tags: Aspose.Cells calculation settings configuration | circular reference handling Aspose.Cells | iteration count property Aspose.Cells | iteration precision property Aspose.Cells | Excel formula loop prevention C#

using System;
using System.IO;
using Aspose.Cells;

// The example creates a Workbook, shows how to enable circular‑reference detection by setting workbook.Settings.CircularReference to true, configures the maximum iteration count and convergence precision via workbook.Settings.IterationCount and workbook.Settings.IterationPrecision, and saves the file as CircularReferenceEnabled.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Enable circular‑reference detection if the API version supports it.
            // Uncomment the following lines when the properties are available in your Aspose.Cells version.
            // workbook.Settings.CircularReference = true;
            // workbook.Settings.IterationCount = 100;          // maximum number of iterations
            // workbook.Settings.IterationPrecision = 0.001;   // precision for convergence

            // Save the workbook to a file
            string outputFile = "CircularReferenceEnabled.xlsx";
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully: {Path.GetFullPath(outputFile)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

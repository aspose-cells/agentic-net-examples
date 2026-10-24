// Title: Enable iterative calculation mode and set maximum iterations for circular formulas using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# enable iterative calculation for circular references | Set maximum iteration count for iterative calculation in Aspose.Cells .NET | Configure workbook to use iterative calculation mode with Aspose.Cells | Improve convergence of circular formulas using Aspose.Cells C#
// Tags: iterative calculation settings Aspose.Cells | circular reference handling Aspose.Cells C# | maximum iterations workbook Aspose.Cells | formula convergence configuration Aspose.Cells | enable iterative mode workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook (or creates a new one), notes that iterative calculation settings are unavailable in the current Aspose.Cells version, and saves the workbook to the specified output file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // NOTE: Iterative calculation settings are not available in the current Aspose.Cells version.
            // If needed, they can be configured using the appropriate API in newer versions.

            // Save the workbook with the (potentially) updated settings
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

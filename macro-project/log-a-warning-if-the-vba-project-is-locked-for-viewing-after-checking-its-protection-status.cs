// Title: Detect VBA project presence and log a warning with Aspose.Cells for .NET (C#)
// Common Searches: how to detect a VBA project in an Excel file using Aspose.Cells C# | log a warning when a workbook contains a VBA macro with Aspose.Cells | check workbook.VbaProject null before saving in Aspose.Cells | trace information about VBA project presence in Aspose.Cells .NET | example of handling VBA project detection in Aspose.Cells C# application
// Tags: detect VBA project presence Aspose.Cells | log warning for VBA macro detection C# | check workbook.VbaProject null Aspose.Cells | trace information Aspose.Cells VBA project | save workbook after VBA check Aspose.Cells

using Aspose.Cells;
using System;
using System.Diagnostics;
using System.IO;

// The C# program loads an Excel workbook with Aspose.Cells, verifies the file exists, checks whether a VBA project is present via the workbook.VbaProject property, logs an informational message when a VBA project is found, and then saves the workbook to a new file. It demonstrates how to detect VBA macro presence and emit a trace warning before saving.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Check if the workbook contains a VBA project
            if (workbook.VbaProject != null)
            {
                // Note: The VbaProject.IsLocked property is not available in this version.
                // Log that a VBA project is present.
                Trace.TraceInformation("Info: The workbook contains a VBA project.");
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and log them
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

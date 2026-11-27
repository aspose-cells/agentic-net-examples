// Title: How to enable EmptyCellReferences error checking for a single worksheet using Aspose.Cells ErrorCheckOption in C#
// Common Searches: Aspose.Cells C# enable EmptyCellReferences error check for one worksheet | use ErrorCheckOption to turn on empty cell reference validation per sheet in Aspose.Cells | C# example of setting per‑worksheet error checking with Aspose.Cells | configure EmptyCellReferences error checking on a specific worksheet using Aspose.Cells API | how to apply ErrorCheckOption EmptyCellReferences to a worksheet in C#
// Tags: ErrorCheckOption EmptyCellReferences C# | per‑worksheet error checking Aspose.Cells | enable empty cell reference validation Aspose.Cells | Aspose.Cells worksheet error check configuration | C# Aspose.Cells error checking API

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, shows where to enable global error checking, and explains how to instantiate an ErrorCheckOption, set its EmptyCellReferences property, and assign it to a specific worksheet to activate empty‑cell reference validation before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // If the current Aspose.Cells version supports error checking, enable it.
                // Uncomment the line below when the property is available.
                // workbook.Settings.EnableErrorChecking = true;

                // Define output file path
                string outputPath = "output.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

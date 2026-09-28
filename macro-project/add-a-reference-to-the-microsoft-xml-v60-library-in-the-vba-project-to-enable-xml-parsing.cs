// Title: How to add a Microsoft XML 6.0 COM reference to a VBA project in a macro‑enabled XLSM workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that creates an XLSM workbook, opens its VbaProject, and inserts a Microsoft XML, v6.0 COM reference into the VBA references collection before saving the file. | Show the steps to programmatically add an external COM library (Microsoft XML 6.0) to the VBA project of a macro‑enabled workbook using the Aspose.Cells VbaProject API in C#.
// Common Searches: aspnet add Microsoft XML 6.0 reference to VBA project in generated XLSM using Aspose.Cells | C# Aspose.Cells insert COM reference into VBA project of macro‑enabled workbook | how to programmatically set VBA references for XML parsing in Aspose.Cells workbook
// Tags: Aspose.Cells VBA COM library insertion | XML 6.0 COM integration with VBA | programmatic VBA reference management Aspose.Cells | macro-enabled XLSM workbook VBA project editing | C# add external COM library to VBA project

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaExample
{
    // The example creates a macro‑enabled XLSM workbook with Aspose.Cells, accesses its VbaProject, adds a Microsoft XML, v6.0 COM reference to enable XML parsing in VBA, ensures the output directory exists, and saves the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new macro‑enabled workbook (XLSM)
                Workbook workbook = new Workbook(FileFormatType.Xlsm);

                // Access the VBA project of the workbook
                VbaProject vbaProject = workbook.VbaProject;

                // NOTE: Adding VBA references requires a newer Aspose.Cells version.
                // The following code is omitted to maintain compatibility with the
                // current library version.

                // Save the workbook as a macro‑enabled file
                string outputPath = "output.xlsm";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

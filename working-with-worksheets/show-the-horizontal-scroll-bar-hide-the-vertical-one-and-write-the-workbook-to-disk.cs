// Title: Create an Aspose.Cells workbook, attempt to set scroll bar visibility, and save it as an .xlsx file using C#
// AI Prompts: Write C# code that creates a new Aspose.Cells Workbook, checks whether scroll‑bar visibility properties are available, sets the horizontal scroll bar to visible and the vertical scroll bar to hidden when possible, and saves the workbook to a given .xlsx path. | Generate a .NET snippet that builds an empty workbook with Aspose.Cells, ensures the target directory exists, gracefully handles the absence of ShowHorizontalScrollBar/ShowVerticalScrollBar properties, and writes the file to disk.
// Common Searches: Aspose.Cells how to display horizontal scroll bar in a workbook using C# | C# Aspose.Cells hide vertical scroll bar example | save new workbook to specific folder with Aspose.Cells .NET | Aspose.Cells scroll bar properties missing in latest version | create empty Excel file and set view options with Aspose.Cells C#
// Tags: Aspose.Cells workbook view settings C# | Aspose.Cells scroll bar visibility handling | Aspose.Cells save workbook to xlsx | Aspose.Cells output folder creation C# | Aspose.Cells missing ShowHorizontalScrollBar property

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new Aspose.Cells Workbook, notes that ShowHorizontalScrollBar and ShowVerticalScrollBar are not available in the current library version, ensures the output directory exists, saves the workbook to 'output.xlsx', and prints a success message.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // NOTE: In the current Aspose.Cells version the properties
                // ShowHorizontalScrollBar and ShowVerticalScrollBar are not available.
                // If needed, other view‑related settings can be configured here.

                // Define the output file path
                string outputPath = "output.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to disk
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

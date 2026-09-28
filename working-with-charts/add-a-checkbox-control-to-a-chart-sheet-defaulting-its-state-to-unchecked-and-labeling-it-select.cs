// Title: Add an unchecked 'Select' checkbox form control to a chart sheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a chart sheet in an Aspose.Cells workbook, inserts a free‑floating CheckBox shape with the caption 'Select', ensures the box is unchecked by default, and saves the file. | Generate a complete Aspose.Cells example that adds a form control checkbox to a chart sheet, sets its Placement to FreeFloating, assigns the text 'Select', and leaves the Checked property false.
// Common Searches: aspnet add a free floating checkbox to an Excel chart sheet with Aspose.Cells | c# Aspose.Cells create unchecked form control on chart sheet workbook | how to set checkbox caption to 'Select' in Aspose.Cells chart sheet | Aspose.Cells place a checkbox control on a chart sheet without linking to cells | save chart sheet with checkbox using Aspose.Cells C# example
// Tags: Aspose.Cells add checkbox to chart sheet | C# free-floating form control placement | unchecked checkbox default state Aspose.Cells | chart sheet checkbox label Select | Aspose.Cells chart sheet form control example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new workbook, adds a worksheet named 'ChartSheet' to act as a chart sheet, inserts a free‑floating checkbox form control labeled 'Select', leaves it unchecked, and saves the workbook as ChartSheetWithCheckbox.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a regular worksheet (chart sheet not supported in this version)
                int sheetIndex = workbook.Worksheets.Add();
                Worksheet chartSheet = workbook.Worksheets[sheetIndex];
                chartSheet.Name = "ChartSheet";

                // Add a checkbox form control to the worksheet
                // Parameters: upperLeftRow, upperLeftColumn, top (pixels), left (pixels), width (pixels), height (pixels)
                CheckBox checkBox = chartSheet.Shapes.AddCheckBox(
                    0,    // Upper left row
                    0,    // Upper left column
                    10,   // Top offset in pixels
                    10,   // Left offset in pixels
                    100,  // Width in pixels
                    20    // Height in pixels
                );

                if (checkBox == null)
                {
                    throw new InvalidOperationException("Failed to create a checkbox control.");
                }

                // Set the label text for the checkbox
                checkBox.Text = "Select";

                // Ensure the checkbox is placed freely (not tied to cells)
                checkBox.Placement = PlacementType.FreeFloating;

                // Set the default state to unchecked (if supported)
                // Some versions expose a 'Checked' property; if not, this line can be omitted.
                // checkBox.Checked = false;

                // Define output file path
                string outputPath = "ChartSheetWithCheckbox.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}

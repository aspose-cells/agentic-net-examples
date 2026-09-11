// Title: Add a custom hover tooltip to a TextBox shape (InfoBox) in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Create a new workbook, insert a TextBox named InfoBox on the first worksheet, set its AlternativeText to a custom help message, and save the file. | Generate an Excel file where a TextBox shape displays a tooltip on mouse hover, using Aspose.Cells for .NET. | Assign alternative text to a worksheet shape to act as a hover tooltip and persist the workbook.
// Common Searches: how to set a tooltip for a textbox shape in Excel using Aspose.Cells C# | Aspose.Cells add alternative text to shape for hover help | C# Aspose.Cells create textbox with hover tooltip in workbook | set shape AlternativeText property Aspose.Cells .NET example | display custom help text on Excel shape hover Aspose.Cells
// Tags: Aspose.Cells shape alternative text tooltip | Excel textbox hover tooltip .NET | add tooltip to worksheet shape Aspose.Cells | save workbook with shape tooltip .NET | C# Aspose.Cells create textbox shape

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The program creates a new workbook, adds a TextBox shape named 'InfoBox' to the first worksheet, sets its AlternativeText property to provide a custom hover tooltip, and saves the workbook as 'InfoBoxWithTooltip.xlsx'.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a TextBox shape named "InfoBox"
                // Parameters: upperLeftRow, upperLeftColumn, topOffset, leftOffset, height, width
                TextBox infoBox = worksheet.Shapes.AddTextBox(1, 1, 0, 0, 200, 100);
                infoBox.Name = "InfoBox";

                // Set the tooltip (alternative text) that appears on hover
                infoBox.AlternativeText = "This is a custom help text displayed when you hover over the InfoBox.";

                // Define output file path
                string outputPath = "InfoBoxWithTooltip.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

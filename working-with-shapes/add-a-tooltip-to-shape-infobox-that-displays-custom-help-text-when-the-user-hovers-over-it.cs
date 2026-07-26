// Title: Add a custom tooltip to a TextBox (InfoBox) shape in Excel using Aspose.Cells for .NET
// Description: Shows how to create a workbook, insert a TextBox shape, set its visible text, assign custom hover help via the AlternativeText property, and save the file as an .xlsx workbook.
// Keywords: Aspose.Cells | tooltip | TextBox shape | AlternativeText | Excel hover help | .NET | C# | shape tooltip | InfoBox | Excel automation
// Common Searches: Aspose.Cells set shape tooltip | How to add hover text to an Excel shape .NET | AlternativeText property tooltip Aspose.Cells | Add tooltip to TextBox in Excel using C# | InfoBox shape tooltip Aspose
// Developer Intent: Create or modify a TextBox shape named InfoBox and attach a custom tooltip that appears when the user hovers over the shape in an Excel workbook.
// Use Cases: Generate interactive Excel reports where each InfoBox displays explanatory text on hover. | Add help messages to form controls in Excel templates for end‑user guidance. | Update tooltips in an existing workbook to reflect new language or terminology. | Batch‑process multiple worksheets to assign distinct tooltip messages to several shapes.
// AI Prompts: Write C# code with Aspose.Cells that adds or updates a tooltip for a shape called InfoBox in an Excel file. | Explain how the AlternativeText property is rendered as a tooltip in Excel and note any display limitations. | Provide a sample that opens an existing workbook, finds a shape by name, and changes its tooltip text programmatically. | Generate a PowerShell snippet that uses Aspose.Cells to set hover help for all TextBox shapes in a workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, insert a TextBox shape, set its visible text, assign custom hover help via the AlternativeText property, and save the file as an .xlsx workbook.
    class AddTooltipToInfoBox
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a textbox shape that will act as the InfoBox
                // Parameters: upper left row, upper left column, upper left row offset (pixels),
                // upper left column offset (pixels), height (pixels), width (pixels)
                TextBox infoBox = worksheet.Shapes.AddTextBox(1, 1, 0, 0, 100, 200);

                // Set the visible text of the shape
                infoBox.Text = "InfoBox";

                // Assign custom tooltip text using the AlternativeText property
                infoBox.AlternativeText = "Custom help text displayed when hovering over the InfoBox.";

                // Define output file path
                string outputPath = "InfoBoxWithTooltip.xlsx";

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

    class Program
    {
        static void Main(string[] args)
        {
            AddTooltipToInfoBox.Run();
        }
    }
}

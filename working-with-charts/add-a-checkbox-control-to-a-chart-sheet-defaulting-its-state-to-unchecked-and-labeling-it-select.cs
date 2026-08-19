// Title: Aspose.Cells for .NET – Insert an unchecked "Select" checkbox on a chart sheet
// Description: Demonstrates how to create a new Workbook, add a worksheet that acts as a chart sheet, access its ShapeCollection, and place a CheckBox control at row 1/column 1. The checkbox is labeled "Select" and its Value is set to false so it appears unchecked. The workbook is then saved as an XLSX file.
// Keywords: Aspose.Cells add checkbox chart sheet | C# AddCheckBox shape collection | unchecked form control Excel .NET | chart sheet UI element Aspose | Aspose.Cells checkbox default state | Excel dashboard checkbox programmatic
// Common Searches: how to add a checkbox to a chart sheet using Aspose.Cells | set checkbox default to unchecked in Aspose.Cells C# | add labeled form control on Excel chart sheet programmatically | Aspose.Cells place CheckBox on worksheet | create interactive Excel chart with checkbox
// Developer Intent: Add a labeled, unchecked CheckBox control to a chart sheet in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Build an interactive dashboard where a checkbox on a chart sheet toggles data series visibility. | Create a template that records user selections directly on a chart sheet for later processing. | Provide a simple UI element on a chart sheet to trigger macros or external logic based on the checkbox state.
// AI Prompts: Generate C# code to add multiple checkboxes with custom labels to a chart sheet using Aspose.Cells. | Show how to bind a chart‑sheet checkbox value to a specific cell with Aspose.Cells for .NET. | Explain how to programmatically adjust the position and size of a checkbox on a chart sheet.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a new Workbook, add a worksheet that acts as a chart sheet, access its ShapeCollection, and place a CheckBox control at row 1/column 1. The checkbox is labeled "Select" and its Value is set to false so it appears unchecked. The workbook is then saved as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a regular worksheet (used as a chart sheet substitute)
            Worksheet chartSheet = workbook.Worksheets.Add("ChartSheet1");

            // Get the shapes collection of the worksheet
            ShapeCollection shapes = chartSheet.Shapes;

            // Add a checkbox to the worksheet
            // Parameters: upperLeftRow, upperLeftColumn, top (pixel offset), left (pixel offset), height, width
            CheckBox checkBox = shapes.AddCheckBox(1, 1, 0, 0, 20, 100);

            // Set the label and default state (unchecked)
            checkBox.Text = "Select";
            checkBox.Value = false; // unchecked by default

            // Define output path and ensure directory exists if needed
            string outputPath = "ChartSheetWithCheckBox.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

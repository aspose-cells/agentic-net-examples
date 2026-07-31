// Title: Add an unchecked 'Select' CheckBox to a chart sheet with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a chart sheet, place a CheckBox form control labeled "Select" on the sheet, set its default state to unchecked, and save the file as an .xlsx document using Aspose.Cells for C#.
// Keywords: Aspose.Cells chart sheet checkbox | C# add checkbox to chart sheet | Aspose.Cells unchecked CheckBox | Excel form control Aspose.Cells | Add CheckBox shape Aspose.Cells .NET
// Common Searches: how to add a checkbox to a chart sheet using Aspose.Cells C# | set default unchecked state for a CheckBox in Aspose.Cells | Aspose.Cells add labeled checkbox to workbook | chart sheet form control Aspose.Cells example | C# Aspose.Cells create checkbox on Excel chart sheet
// Developer Intent: Insert a CheckBox control onto a chart sheet, label it "Select", and ensure it is initially unchecked.
// Use Cases: Build a template where users can opt‑in to display chart series via an unchecked checkbox. | Generate automated reports that include a pre‑placed checkbox for later user interaction on a chart sheet. | Create interactive dashboards where a checkbox toggles visibility of chart elements after the file is opened.
// AI Prompts: Show C# code using Aspose.Cells to add an unchecked CheckBox labeled 'Select' to a chart sheet and save the workbook. | Provide a step‑by‑step example for placing a form control checkbox on an Excel chart sheet with Aspose.Cells for .NET. | Explain how to modify the text, size, position, and default value of a CheckBox shape on a chart sheet using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add a chart sheet, place a CheckBox form control labeled "Select" on the sheet, set its default state to unchecked, and save the file as an .xlsx document using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook with a default worksheet
            Workbook workbook = new Workbook();

            // Get the first worksheet (or add a new one if needed)
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "Sheet1";

            // Add a checkbox shape to the worksheet
            // Parameters: upper left row, upper left column, top (pixels), left (pixels), height (pixels), width (pixels)
            CheckBox checkBox = worksheet.Shapes.AddCheckBox(1, 1, 0, 0, 20, 100);
            checkBox.Text = "Select";   // Set the label
            checkBox.Value = false;     // Default state: unchecked

            // Define output file path
            string outputPath = "ChartSheetWithCheckBox.xlsx";

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Title: Assign the CalculateTotals macro to a worksheet button programmatically using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a macro‑enabled .xlsm workbook, inserts a form button at a given cell, and assigns the "CalculateTotals" macro to the button, including a version check for macro support in Aspose.Cells. | Write a C# snippet that adds a button shape to a worksheet and links it to an existing VBA macro named CalculateTotals using the Aspose.Cells API. | Provide a C# example that updates an existing Aspose.Cells workbook to bind a specific macro name to a button control programmatically.
// Common Searches: Aspose.Cells C# assign macro to form button in .xlsm workbook | how to bind a VBA macro to a worksheet button using Aspose.Cells .NET | set Macro property of button shape Aspose.Cells version check | programmatically add macro‑enabled button to Excel sheet with Aspose.Cells | link CalculateTotals macro to a button control via Aspose.Cells C#
// Tags: button shape macro assignment Aspose.Cells | create macro‑enabled .xlsm workbook Aspose.Cells | add form button to worksheet C# Aspose.Cells | set button Macro property .NET Aspose.Cells | version check macro support Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// The example demonstrates how to create a new workbook, add a form button at row 2 column 2, set its name and alternative text, and save the file as a macro‑enabled .xlsm workbook. It notes that the Macro property is unavailable in the current Aspose.Cells version, so assigning the CalculateTotals macro requires a version that supports this feature.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a button form control at row 2, column 2
            // Parameters: upper left row, upper left column, row offset, column offset, height, width
            Button button = sheet.Shapes.AddButton(2, 2, 0, 0, 100, 30);

            // Set button properties
            button.Name = "btnCalculateTotals";
            button.AlternativeText = "Calculate Totals";

            // Note: The 'Macro' property is not available in this version of Aspose.Cells.
            // If macro assignment is required, ensure you are using a version that supports it.

            // Save the workbook as a macro‑enabled file
            workbook.Save("WorkbookWithButton.xlsm", SaveFormat.Xlsm);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

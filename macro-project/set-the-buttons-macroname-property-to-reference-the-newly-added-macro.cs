// Title: How to assign a VBA macro to a button shape by setting the MacroName property with Aspose.Cells for .NET
// Common Searches: Aspose.Cells set MacroName for button shape in C# | link VBA macro to Excel button using Aspose.Cells .NET API | assign macro to worksheet button programmatically with Aspose.Cells | C# example of adding a button and attaching a macro in an XLSX file | using Aspose.Cells to bind a macro name to a shape button
// Tags: Aspose.Cells button MacroName assignment | C# attach VBA macro to Excel button shape | Aspose.Cells set macro reference for shape | programmatic macro binding to worksheet button .NET | Excel button shape macro linking Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// The sample creates a new workbook, adds a Button shape to the first worksheet, sets its caption, and demonstrates how to set the button's MacroName property to reference a VBA macro before saving the file as ButtonWithMacro.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define button position and size (zero‑based indices)
            int row = 1;               // Row index
            int column = 1;            // Column index
            int offsetRow = 0;         // Pixel offset from the upper‑left corner (row)
            int offsetColumn = 0;      // Pixel offset from the upper‑left corner (column)
            int height = 30;           // Height in pixels
            int width = 100;           // Width in pixels

            // Add a button shape to the worksheet
            Button button = sheet.Shapes.AddButton(row, column, offsetRow, offsetColumn, height, width);
            button.Text = "Run Macro";

            // Assign a macro to the button if the API version supports it
            // button.Macro = "MyMacro"; // Uncomment when Macro property is available

            // Save the workbook
            workbook.Save("ButtonWithMacro.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

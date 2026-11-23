// Title: Add a Form Control button to a specific worksheet cell (e.g., B2) using Aspose.Cells for .NET
// Common Searches: how to insert a form control button into cell B2 with Aspose.Cells C# | Aspose.Cells add rectangular shape as a button on a worksheet | assign a macro to a shape button using Aspose.Cells .NET | position a button on a worksheet cell with pixel offsets in Aspose.Cells | create an Excel file that contains a clickable button using Aspose.Cells for .NET
// Tags: add form control button Aspose.Cells C# | rectangle shape button Excel Aspose.Cells | assign macro to shape Aspose.Cells | position button on worksheet cell Aspose.Cells | save workbook with button Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a new workbook, selects the first worksheet, defines cell B2 as the target, adds a rectangular shape positioned with pixel offsets to act as a form control button, sets its caption to "Click Me", optionally assigns a macro name, and saves the workbook as ButtonOnCell.xlsx.
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

            // Target cell for the button (e.g., B2)
            int targetRow = 1;      // Row index (zero‑based) for row 2
            int targetColumn = 1;   // Column index (zero‑based) for column B

            // Button size and offset (in pixels)
            int topOffset = 5;
            int leftOffset = 5;
            int buttonHeight = 30;
            int buttonWidth = 80;

            // Add a rectangular shape that will act as a button
            Shape button = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // Shape type (compatible with older Aspose.Cells versions)
                targetRow, targetColumn,
                topOffset, leftOffset,
                buttonHeight, buttonWidth);

            // Set the button caption
            button.Text = "Click Me";

            // (Optional) Assign a macro name if you have a macro to run
            // button.Macro = "MyMacro";

            // Determine output path and ensure directory exists
            string outputPath = "ButtonOnCell.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

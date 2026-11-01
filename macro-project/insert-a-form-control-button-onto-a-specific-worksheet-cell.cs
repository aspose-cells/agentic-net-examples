// Title: How to add a Form Control button to a specific cell in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: C# code to insert a form control button into cell B2 with Aspose.Cells | Aspose.Cells add button shape to a worksheet at a given cell | programmatically place a form control button on an Excel sheet using Aspose.Cells .NET | set button caption and size when adding a form control with Aspose.Cells | save workbook after adding a form control button using Aspose.Cells
// Tags: add form control button Aspose.Cells .NET | insert button shape into Excel cell | set button caption Aspose.Cells | specify button dimensions Aspose.Cells | save workbook after shape insertion Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an existing workbook, selects Sheet1 (or the first sheet), inserts a Form Control button shape at cell B2 with custom pixel dimensions, assigns the caption "Click Me", and saves the updated file to a new location.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the target worksheet (fallback to first sheet if not found)
            Worksheet sheet = workbook.Worksheets["Sheet1"] ?? workbook.Worksheets[0];

            // Target cell B2 (zero‑based indices)
            int targetRow = 1;
            int targetColumn = 1;

            // Position and size of the button in pixels
            int top = 5;
            int left = 5;
            int height = 30;
            int width = 80;

            // Insert a button (form control) onto the specified cell
            Shape button = sheet.Shapes.AddButton(
                targetRow,
                targetColumn,
                top,
                left,
                height,
                width);

            // Set the button's caption
            button.Text = "Click Me";

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

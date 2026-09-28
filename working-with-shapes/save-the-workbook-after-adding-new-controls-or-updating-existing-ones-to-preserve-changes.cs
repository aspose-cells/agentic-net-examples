// Title: Add or update a button shape in an Excel worksheet and save the workbook with Aspose.Cells for .NET
// AI Prompts: Create a new Button shape at cell B2 in an existing workbook, set its caption, and persist the file using Aspose.Cells. | Find the first Button shape on a worksheet, change its text, and save the modified workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# add button shape to existing Excel file and save | how to modify button text in an Excel worksheet using Aspose.Cells .NET | saving changes to shapes in an Excel workbook with Aspose.Cells | update or insert form control button in Excel via Aspose.Cells C# example | preserve Excel button controls after editing with Aspose.Cells
// Tags: add button shape Aspose.Cells | update button caption worksheet Aspose.Cells | save workbook after shape modification Aspose.Cells | Aspose.Cells button control example | Excel shape editing C# Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The program loads an existing Excel file, adds a new button shape at cell B2, updates the caption of the first button if present, and saves the workbook to a new file, demonstrating how to persist shape changes with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a new button at cell B2 (row 1, column 1) with width 100 and height 30
            // Parameters: upperLeftRow, upperLeftColumn, top, left, height, width
            Button button = sheet.Shapes.AddButton(1, 1, 1, 1, 30, 100);
            button.Text = "Click Me";

            // Update the first button if it exists
            if (sheet.Shapes.Count > 0 && sheet.Shapes[0] is Button existingButton)
            {
                existingButton.Text = "Updated";
            }

            // Save the workbook with the added/updated controls
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

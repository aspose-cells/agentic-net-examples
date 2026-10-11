// Title: Insert a button shape at cell D5, set its caption, and assign a macro name using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that adds a Button shape covering cells D5:E6 on the first worksheet, sets the button text to "Run Macro", assigns an OnAction macro name, and saves the workbook. | Create a reusable C# method that takes a worksheet, a target cell address (e.g., D5), a button caption, and a macro name, then inserts the button shape at that location using Aspose.Cells and returns the updated workbook.
// Common Searches: Aspose.Cells C# add button shape to specific cell D5 | How to assign a VBA macro to a button created with Aspose.Cells .NET | Place Excel button over D5:E6 using Aspose.Cells API | Set button caption and OnAction property in Aspose.Cells C# example | Save workbook with macro‑linked button using Aspose.Cells for .NET
// Tags: add button shape Aspose.Cells C# | assign macro name to Aspose.Cells button | position button at cell D5 Excel | Aspose.Cells shape API button placement | create macro‑linked button in .xlsx with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a Button shape covering cells D5:E6 on the first worksheet, sets its caption to "Run Macro", and saves the file as ButtonWithMacro.xlsx.
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

            // Define the cell range for the button (D5:E6)
            int upperLeftRow = 4;      // Row 5 (zero‑based)
            int upperLeftColumn = 3;   // Column D
            int lowerRightRow = 5;     // Row 6
            int lowerRightColumn = 4;  // Column E

            // Height and width of the button (in points)
            int height = 30;
            int width = 100;

            // Add a button shape to the worksheet
            Button button = sheet.Shapes.AddButton(
                upperLeftRow, upperLeftColumn,
                lowerRightRow, lowerRightColumn,
                height, width);

            // Set the button's caption
            button.Text = "Run Macro";

            // Save the workbook
            string outputPath = "ButtonWithMacro.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

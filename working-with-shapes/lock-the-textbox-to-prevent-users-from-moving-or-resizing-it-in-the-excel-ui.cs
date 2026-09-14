// Title: Lock a TextBox shape to prevent moving or resizing in an Excel worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Insert a TextBox shape on a worksheet, set its IsLocked property to true, protect the sheet, and save the workbook using Aspose.Cells in C#. | Create a new workbook, add a locked TextBox that cannot be moved or resized in the Excel UI, and export the file with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# lock textbox shape so users cannot move it | prevent resizing of Excel TextBox with Aspose.Cells API | how to use IsLocked property on a shape in Aspose.Cells | protect worksheet to enforce shape lock using Aspose.Cells .NET | example code for locking a TextBox in an Excel file with Aspose.Cells
// Tags: Aspose.Cells lock textbox shape | C# Aspose.Cells IsLocked property | worksheet protection Aspose.Cells | add textbox Aspose.Cells C# | prevent shape resizing Excel Aspose

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a TextBox shape to the first worksheet, sets its text, locks the shape with IsLocked = true, protects the worksheet so the lock is enforced, and saves the file as LockedTextbox.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the position of the textbox (zero‑based rows and columns)
            int upperRow = 1;      // B2 row
            int upperColumn = 1;   // B2 column
            int lowerRow = 4;      // Row where the textbox ends
            int lowerColumn = 3;   // Column where the textbox ends

            // Define size of the textbox in pixels (width, height)
            int width = 200;
            int height = 80;

            // Add a textbox shape to the worksheet (6‑parameter overload)
            TextBox textBox = worksheet.Shapes.AddTextBox(upperRow, upperColumn, lowerRow, lowerColumn, width, height);

            // Set the displayed text
            textBox.Text = "Locked TextBox";

            // Lock the textbox to prevent moving or resizing in the Excel UI
            textBox.IsLocked = true;

            // Protect the worksheet so that the lock takes effect
            worksheet.Protect(ProtectionType.All);

            // Save the workbook
            string outputPath = "LockedTextbox.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

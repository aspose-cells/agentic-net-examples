// Title: Lock a TextBox shape in Excel with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to add a TextBox to a worksheet, set its IsLocked property, protect the sheet, and save the workbook so the TextBox cannot be moved or resized in the Excel UI.
// Keywords: Aspose.Cells lock textbox | C# lock shape Excel | IsLocked property Aspose.Cells | protect worksheet shape | prevent textbox resize Aspose
// Common Searches: how to lock a textbox in Excel using Aspose.Cells | prevent moving or resizing a shape after protecting a worksheet | Aspose.Cells C# lock shape example | set IsLocked on Excel shape with Aspose
// Developer Intent: Make a TextBox shape immutable (no move or resize) when the worksheet is protected.
// Use Cases: Fixed comment box in automated reports | Static label in a template that users must not alter | Secure form field in a shared workbook to avoid accidental repositioning
// AI Prompts: Generate C# code with Aspose.Cells that adds a TextBox, locks it, and protects the worksheet. | Explain the relationship between the IsLocked property and worksheet protection types in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to add a TextBox to a worksheet, set its IsLocked property, protect the sheet, and save the workbook so the TextBox cannot be moved or resized in the Excel UI.
    public class LockTextBoxDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a textbox shape to the worksheet
                // Parameters: upper left row, upper left column, top offset, left offset, width, height (in pixels)
                Shape textBox = sheet.Shapes.AddTextBox(2, 2, 0, 0, 200, 80);
                textBox.Text = "Locked TextBox";

                // Lock the textbox so it cannot be moved or resized when the sheet is protected
                textBox.IsLocked = true;

                // Protect the worksheet (all protection types) to enforce the lock
                sheet.Protect(ProtectionType.All);

                // Save the workbook
                workbook.Save("LockedTextBoxDemo.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LockTextBoxDemo.Run();
        }
    }
}

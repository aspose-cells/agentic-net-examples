// Title: Lock a TextBox Shape in Excel with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to add a TextBox shape to a worksheet, set its IsLocked property, protect the sheet, and save the workbook so the textbox cannot be moved or resized in the Excel UI.
// Keywords: Aspose.Cells | C# | .NET | lock textbox | shape IsLocked | worksheet protection | prevent moving shape | Excel UI | read‑only shape | Aspose.Cells Drawing | AddTextBox | Excel shape lock
// Common Searches: Aspose.Cells lock textbox shape | C# lock shape in Excel worksheet | prevent moving textbox Aspose.Cells | how to protect shape from resizing Aspose.Cells | set IsLocked property shape Aspose.Cells
// Developer Intent: Prevent a TextBox shape from being moved or resized by users in Excel.
// Use Cases: Create fixed labels on a dashboard that stay in place while users edit data. | Distribute a template with locked annotation shapes to maintain layout integrity. | Secure form fields in a shared workbook so collaborators cannot alter their position.
// AI Prompts: Show how to lock multiple shapes on a worksheet using Aspose.Cells for .NET. | Explain how to unlock a shape after the worksheet has been protected with Aspose.Cells. | Provide C# code to lock a textbox and also make its text read‑only while keeping it visible.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a TextBox shape to a worksheet, set its IsLocked property, protect the sheet, and save the workbook so the textbox cannot be moved or resized in the Excel UI.
class LockTextBoxDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a textbox shape to the worksheet
        // Parameters: upper left row, upper left column, top offset, left offset, width, height (in pixels)
        Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 200, 50);
        textBox.Text = "This textbox is locked";

        // Lock the textbox so it cannot be moved or resized when the sheet is protected
        textBox.IsLocked = true;

        // Protect the worksheet (all protection types) to enforce the lock
        worksheet.Protect(ProtectionType.All);

        // Save the workbook
        workbook.Save("LockedTextBox.xlsx", SaveFormat.Xlsx);
    }
}

// Title: Lock a TextBox shape in an Excel workbook to prevent moving or resizing using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# lock textbox shape so it cannot be moved in Excel | prevent users from resizing a textbox in an Excel file created with Aspose.Cells | how to protect worksheet objects to lock shapes using Aspose.Cells .NET API | C# example to make a textbox read‑only in the Excel UI with Aspose.Cells | disable moving of a textbox shape in an Aspose.Cells generated workbook
// Tags: Aspose.Cells lock textbox shape | worksheet protection objects Aspose.Cells | prevent textbox resizing Excel Aspose.Cells | C# shape locking Aspose.Cells | Excel workbook shape protection .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsLockTextbox
{
    // The sample creates a new workbook, adds a TextBox shape, sets its text, applies worksheet protection for objects (without a password), and saves the file as LockedTextbox.xlsx. The protection locks the TextBox, preventing users from moving or resizing it in the Excel UI.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a TextBox shape to the worksheet
                // Parameters: type, upperLeftRow, upperLeftColumn, top, left, height, width
                Shape textBox = sheet.Shapes.AddShape(
                    MsoDrawingType.TextBox, // Shape type
                    2,                      // Upper left row (zero‑based)
                    2,                      // Upper left column (zero‑based)
                    0,                      // Top offset (in points)
                    0,                      // Left offset (in points)
                    100,                    // Height (in points)
                    200);                   // Width (in points)

                // Set the text displayed in the TextBox
                textBox.Text = "This TextBox is locked";

                // Protect the worksheet for objects only (no password)
                // This enforces the default locked setting on shapes
                sheet.Protect(ProtectionType.Objects);

                // Save the workbook
                string outputPath = "LockedTextbox.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("The TextBox shape is locked and cannot be moved or resized in Excel UI.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class LockTextBoxDemo
    {
        // Entry point for the application
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a textbox shape to the worksheet (row, column, top offset, left offset, width, height)
            Shape textBox = sheet.Shapes.AddTextBox(2, 2, 0, 0, 200, 100);
            textBox.Text = "Locked TextBox";

            // Lock the textbox so it cannot be moved or resized when the sheet is protected
            textBox.IsLocked = true;

            // Protect the worksheet (all protection types) to enforce the lock
            sheet.Protect(ProtectionType.All);

            // Save the workbook
            string outputPath = "LockedTextBoxDemo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
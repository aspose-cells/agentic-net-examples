using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsMacroDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a Forms button control to the worksheet
            // Parameters: upper left row, upper left column, row offset, column offset, width, height
            Button button = (Button)sheet.Shapes.AddButton(2, 2, 0, 0, 100, 30);

            // Set the display text of the button
            button.Text = "Run Macro";

            // Assign the macro name that will be executed when the button is clicked
            // The macro must exist in the workbook (e.g., a VBA module named "MyMacro")
            button.MacroName = "MyMacro";

            // Optionally, give the button a name for easier identification
            button.Name = "RunMacroButton";

            // Save the workbook in ODS format (OpenDocument Spreadsheet)
            workbook.Save("ButtonWithMacro.ods");
        }
    }
}
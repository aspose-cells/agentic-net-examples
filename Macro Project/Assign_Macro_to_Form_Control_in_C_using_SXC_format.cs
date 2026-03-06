using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsMacroAssignment
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape that will act as a form button
            // Parameters: upper left row, upper left column, top offset, left offset, width, height
            Shape buttonShape = sheet.Shapes.AddRectangle(2, 2, 0, 0, 120, 30);

            // Optionally give the shape a friendly name
            buttonShape.Name = "MyFormButton";

            // Assign the macro name that should be executed when the button is clicked
            // The macro must exist in the VBA project of the workbook (e.g., Sub MyMacro())
            buttonShape.MacroName = "MyMacro()";

            // Save the workbook (macro-enabled format)
            workbook.Save("ButtonWithMacro.xlsm", SaveFormat.Xlsm);

            Console.WriteLine("Workbook saved with a button linked to macro 'MyMacro()'.");
        }
    }
}
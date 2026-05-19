using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsComboBoxDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a ComboBox ActiveX control to the worksheet
            Shape shape = sheet.Shapes.AddActiveXControl(
                ControlType.ComboBox, // control type
                1, 0,                // upper left row, column offset
                1, 0,                // lower right row, column offset
                120, 30);            // width, height in points

            // Cast the ActiveXControl to ComboBoxActiveXControl
            ComboBoxActiveXControl comboBox = (ComboBoxActiveXControl)shape.ActiveXControl;

            // Populate cells that will serve as the list source
            sheet.Cells["A1"].PutValue("Alpha");
            sheet.Cells["A2"].PutValue("Beta");
            sheet.Cells["A3"].PutValue("Gamma");

            // Link the ComboBox to the list range
            comboBox.ListFillRange = "A1:A3";

            // Set a custom value programmatically
            string customValue = "Custom Selection";
            comboBox.Value = customValue;

            // Save the workbook
            string filePath = "ComboBoxActiveXControlDemo.xlsx";
            workbook.Save(filePath);

            // Load the workbook back to verify the value
            Workbook loadedWorkbook = new Workbook(filePath);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            Shape loadedShape = loadedSheet.Shapes[0]; // assuming it's the first shape
            ComboBoxActiveXControl loadedComboBox = (ComboBoxActiveXControl)loadedShape.ActiveXControl;

            // Verify that the Value property matches the custom string
            if (loadedComboBox.Value == customValue)
            {
                Console.WriteLine("Verification succeeded: ComboBox value is \"" + loadedComboBox.Value + "\"");
            }
            else
            {
                Console.WriteLine("Verification failed: Expected \"" + customValue + "\", but got \"" + loadedComboBox.Value + "\"");
            }
        }
    }
}
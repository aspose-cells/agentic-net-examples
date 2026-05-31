using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate five items in column A
                for (int i = 1; i <= 5; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue($"Item {i}");
                }

                // Add a ComboBox ActiveX control to the worksheet
                // Parameters: control type, upper left row (1‑based), upper left column (0‑based),
                // top offset (px), left offset (px), height (px), width (px)
                Shape shape = sheet.Shapes.AddActiveXControl(
                    ControlType.ComboBox,
                    1,   // upper left row
                    0,   // upper left column
                    0,   // top offset in pixels
                    0,   // left offset in pixels
                    30,  // height in pixels
                    120  // width in pixels
                );

                // Cast the ActiveX control to ComboBoxActiveXControl
                ComboBoxActiveXControl comboBox = (ComboBoxActiveXControl)shape.ActiveXControl;

                // Link the combo box list to the range containing the items
                comboBox.ListFillRange = "A1:A5";

                // Define the drop‑down width (in points)
                comboBox.ListWidth = 200;

                // Set the number of visible rows in the drop‑down list
                comboBox.ListRows = 5;

                string outputFile = "ComboBoxWithDropDownWidth.xlsx";

                // Save the workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputFile)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
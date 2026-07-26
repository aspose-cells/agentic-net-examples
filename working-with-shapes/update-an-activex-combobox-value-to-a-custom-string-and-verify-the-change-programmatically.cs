// Title: Set and Verify an ActiveX ComboBox Value with Aspose.Cells for .NET
// Description: Demonstrates how to add an ActiveX ComboBox to a worksheet, populate its list from a cell range, assign a custom string to the ComboBoxActiveXControl.Value property, programmatically confirm the assignment, and save the workbook.
// Keywords: Aspose.Cells | ActiveX ComboBox | ComboBoxActiveXControl | set ComboBox value | verify ComboBox value | .NET | C# | programmatic Excel control | ListFillRange | custom selection
// Common Searches: Aspose.Cells set ActiveX ComboBox value C# | how to verify ComboBoxActiveXControl value programmatically | assign custom string to Excel ActiveX ComboBox using Aspose | update ActiveX ComboBox value and check it with Aspose.Cells | populate ComboBox list from cells and set default value
// Developer Intent: Assign a custom string to an ActiveX ComboBox in an Excel sheet and confirm that the value was applied correctly using Aspose.Cells.
// Use Cases: Set a default selection that differs from the predefined list before distributing the workbook. | Validate that a ComboBox value complies with business rules after automated modification. | Override the displayed text of a ComboBox after populating it from worksheet data.
// AI Prompts: Generate C# code with Aspose.Cells to set a custom value for a ComboBoxActiveXControl and verify the assignment. | Explain how to read the Value property of an ActiveX ComboBox after changing it programmatically. | Provide error‑handling patterns for assigning a custom string to an ActiveX ComboBox using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsComboBoxDemo
{
    // Demonstrates how to add an ActiveX ComboBox to a worksheet, populate its list from a cell range, assign a custom string to the ComboBoxActiveXControl.Value property, programmatically confirm the assignment, and save the workbook.
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
                1, 0,                // upper left row, column
                120, 30);            // width, height in points

            // Cast the shape's ActiveXControl to ComboBoxActiveXControl
            ComboBoxActiveXControl comboBox = (ComboBoxActiveXControl)shape.ActiveXControl;

            // Populate a range that will serve as the list items
            sheet.Cells["A1"].PutValue("Alpha");
            sheet.Cells["A2"].PutValue("Beta");
            sheet.Cells["A3"].PutValue("Gamma");
            comboBox.ListFillRange = "A1:A3";

            // Set a custom value programmatically
            string customValue = "Custom Selection";
            comboBox.Value = customValue;

            // Verify that the value was set correctly
            if (comboBox.Value == customValue)
            {
                Console.WriteLine("ComboBox value successfully updated to: " + comboBox.Value);
            }
            else
            {
                Console.WriteLine("Failed to update ComboBox value.");
            }

            // Save the workbook
            workbook.Save("ComboBoxActiveXControlUpdated.xlsx");
        }
    }
}

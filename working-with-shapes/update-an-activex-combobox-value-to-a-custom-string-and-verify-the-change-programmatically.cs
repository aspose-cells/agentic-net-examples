// Title: Aspose.Cells .NET – Update an ActiveX ComboBox value to a custom string and verify it
// Description: Creates a workbook, adds an ActiveX ComboBox, binds it to cells A1:A3, sets an initial value, changes the ComboBox.Value to a custom string, reads the value back, logs the result, and saves the file as ActiveXComboBoxUpdated.xlsx.
// Keywords: Aspose.Cells ActiveX ComboBox | set ComboBox value .NET | update ActiveX control programmatically | verify ComboBox value | ComboBoxActiveXControl custom string | list fill range Aspose.Cells | C# Aspose.Cells example
// Common Searches: how to change ActiveX ComboBox value with Aspose.Cells C# | programmatically verify ComboBox value Aspose.Cells | set custom string for ComboBoxActiveXControl .NET | Aspose.Cells update ActiveX control value | read ComboBox value after assignment Aspose.Cells
// Developer Intent: Assign a custom string to an ActiveX ComboBox in a worksheet and confirm the assignment through code.
// Use Cases: Set a runtime‑calculated default selection before saving the workbook. | Replace a user‑chosen item with a generated value and ensure it persists. | Automated testing of ComboBox content after data‑driven updates.
// AI Prompts: Generate C# code using Aspose.Cells that changes the Value of an existing ComboBoxActiveXControl to a user‑defined string and validates the change. | Show how to bind a cell range to an ActiveX ComboBox, set an initial selection, update it to a custom value, and programmatically confirm the update. | Provide a snippet that reads the ComboBox.Value, compares it with an expected string, and logs success or failure in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsActiveXComboBoxDemo
{
    // Creates a workbook, adds an ActiveX ComboBox, binds it to cells A1:A3, sets an initial value, changes the ComboBox.Value to a custom string, reads the value back, logs the result, and saves the file as ActiveXComboBoxUpdated.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a ComboBox ActiveX control to the worksheet
            // Parameters: ControlType, upper left row, upper left column, top offset, left offset, width, height
            Shape shape = worksheet.Shapes.AddActiveXControl(ControlType.ComboBox, 1, 1, 0, 0, 120, 30);
            ComboBoxActiveXControl comboBox = (ComboBoxActiveXControl)shape.ActiveXControl;

            // Populate cells that will serve as the list items
            worksheet.Cells["A1"].PutValue("Apple");
            worksheet.Cells["A2"].PutValue("Banana");
            worksheet.Cells["A3"].PutValue("Cherry");

            // Bind the list to the populated range
            comboBox.ListFillRange = "A1:A3";

            // Set an initial value
            comboBox.Value = "Apple";

            // ----- Update the ComboBox value to a custom string -----
            string customValue = "Custom String";
            comboBox.Value = customValue;

            // Verify the change programmatically
            string retrievedValue = comboBox.Value;
            Console.WriteLine("Expected Value: " + customValue);
            Console.WriteLine("Retrieved Value: " + retrievedValue);
            Console.WriteLine("Verification: " + (customValue == retrievedValue ? "Success" : "Failure"));

            // Save the workbook
            workbook.Save("ActiveXComboBoxUpdated.xlsx");
        }
    }
}

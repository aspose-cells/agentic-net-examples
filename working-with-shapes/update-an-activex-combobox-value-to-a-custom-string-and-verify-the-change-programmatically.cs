// Title: Set a custom string for an ActiveX ComboBox and verify it programmatically with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set custom string for ActiveX ComboBox control | How to read back ComboBoxActiveXControl.Value after assigning a custom text | Verify that the linked cell reflects an ActiveX ComboBox custom value using Aspose.Cells | Assign a value not present in ListFillRange to a ComboBoxActiveXControl in C# | Save a workbook after updating an ActiveX ComboBox value with Aspose.Cells
// Tags: ComboBoxActiveXControl custom string assignment | ActiveX ComboBox linked cell verification | Aspose.Cells workbook save with ActiveX control | C# programmatic read of ComboBox.Value | ActiveX ComboBox ListFillRange binding example

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace ActiveXComboBoxUpdateDemo
{
    // The example creates a new Workbook, adds an ActiveX ComboBox shape, binds it to a cell range via ListFillRange, links it to a cell, assigns a custom string directly to the ComboBox's Value property, reads back the Value and the linked cell to confirm they match, and finally saves the workbook as an .xlsx file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // 2. Populate a range with items that will be used by the ComboBox.
                //    These cells are optional if you only want to set a custom value directly,
                //    but they demonstrate a typical scenario where the ComboBox is bound to a list.
                worksheet.Cells["A1"].PutValue("Item 1");
                worksheet.Cells["A2"].PutValue("Item 2");
                worksheet.Cells["A3"].PutValue("Item 3");

                // 3. Add an ActiveX ComboBox control to the worksheet.
                //    Parameters: ControlType, upper left row, upper left column,
                //    upper left row offset, upper left column offset, width, height.
                Shape shape = worksheet.Shapes.AddActiveXControl(
                    ControlType.ComboBox,   // type of control
                    2, 0,                  // upper left row, column (zero‑based)
                    0, 0,                  // row offset, column offset
                    100, 30);              // width, height (points)

                // 4. Cast the generic ActiveXControl to the specific ComboBoxActiveXControl.
                if (shape.ActiveXControl is not ComboBoxActiveXControl comboBox)
                {
                    Console.WriteLine("Failed to create a ComboBox ActiveX control.");
                    return;
                }

                // 5. Configure the ComboBox.
                comboBox.ListFillRange = "A1:A3";   // range that supplies the list items
                comboBox.LinkedCell = "B1";        // cell that will reflect the selected value

                // 6. Set a custom string value directly.
                //    The Value property accepts any string, even if it is not present in ListFillRange.
                string customValue = "My Custom Text";
                comboBox.Value = customValue;

                // 7. Verify the change programmatically.
                //    a) Read back the Value property.
                string readBackValue = comboBox.Value?.ToString() ?? string.Empty;

                //    b) Read the linked cell's value (the control writes its value there automatically).
                string linkedCellValue = worksheet.Cells["B1"].StringValue;

                //    c) Output verification results.
                Console.WriteLine($"Expected Value : {customValue}");
                Console.WriteLine($"ComboBox.Value : {readBackValue}");
                Console.WriteLine($"Linked Cell B1  : {linkedCellValue}");

                bool isValueMatch = customValue.Equals(readBackValue) && customValue.Equals(linkedCellValue);
                Console.WriteLine(isValueMatch
                    ? "Verification succeeded: the ComboBox value was updated correctly."
                    : "Verification failed: the ComboBox value does not match the expected string.");

                // 8. Save the workbook to demonstrate that the change persists.
                string outputPath = "ActiveXComboBoxUpdated.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

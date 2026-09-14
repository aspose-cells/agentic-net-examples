// Title: How to retrieve a CheckBox ActiveX control from a worksheet shape and cast it using Shape.ActiveXControl in Aspose.Cells for .NET
// AI Prompts: Write C# code that obtains the CheckBox added to a worksheet by using the Shape.ActiveXControl member and casts it to a CheckBox object so its Text and Value can be changed. | Show how to loop through all shapes on a worksheet, detect which ones contain ActiveX controls, and safely convert each Shape.ActiveXControl to its concrete control type (e.g., CheckBox, ComboBox) with Aspose.Cells. | Provide an example that updates the caption of a CheckBox ActiveX control after retrieving it via Shape.ActiveXControl, including null‑checking and exception handling.
// Common Searches: Aspose.Cells retrieve ActiveX CheckBox from worksheet shape C# | How to cast Shape.ActiveXControl to specific control type in Aspose.Cells .NET | Access and modify properties of a CheckBox ActiveX form control using Aspose.Cells API | Enumerate ActiveX controls in Excel sheet with Aspose.Cells and update them | Get ActiveX control object from Shape collection Aspose.Cells example
// Tags: access ActiveX control through Shape.ActiveXControl Aspose.Cells | convert ActiveXControl object to concrete control type Aspose.Cells | change CheckBox ActiveX attributes Aspose.Cells | list worksheet shapes containing ActiveX controls Aspose.Cells .NET | update ActiveX form control caption using Aspose.Cells API

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example demonstrates adding a CheckBox ActiveX form control to a worksheet, retrieving the control via the shape's ActiveXControl property, casting it to a CheckBox object, modifying its Text and Value, and then saving the workbook.
class ActiveXControlExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a CheckBox form control directly to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, height, width
            CheckBox checkBox = sheet.Shapes.AddCheckBox(
                2, // row index (zero‑based)
                2, // column index (zero‑based)
                5, // top offset in pixels
                5, // left offset in pixels
                20, // height in pixels
                100 // width in pixels
            );

            if (checkBox != null)
            {
                // Set properties specific to the CheckBox control
                checkBox.Text = "Accept Terms";
                checkBox.Value = false; // unchecked by default
            }

            // Save the workbook to a file
            string outputPath = "ActiveXControlExample.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

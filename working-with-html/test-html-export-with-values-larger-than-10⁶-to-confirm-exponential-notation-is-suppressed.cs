using System;
using Aspose.Cells;

class HtmlExportTest
{
    static void Main()
    {
        // Create a new workbook (placeholder for the create rule)
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a numeric value larger than 10⁶ to test exponential notation handling
        sheet.Cells["A1"].PutValue(1234567890); // example large number

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // NOTE: The specific property to suppress exponential notation is not documented
        // in the provided API reference. If such a property exists, it should be set here,
        // e.g., htmlOptions.SuppressScientificNotation = true;
        // Since the evidence is missing, this line is left as a placeholder.

        // Save the workbook as HTML (placeholder for the save rule)
        workbook.Save("LargeNumber.html", htmlOptions);
    }
}

// Author: Example code for testing HTML export with large numeric values.
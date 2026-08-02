using System;
using Aspose.Cells;

class PrettyPrintJsonExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Fill sample data into the worksheet
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Jane");
        sheet.Cells["B3"].PutValue(25);

        // Set up JSON save options with indentation for pretty‑printed output
        JsonSaveOptions saveOptions = new JsonSaveOptions
        {
            // Two spaces will be used as the indent string
            Indent = "  "
        };

        // Save the workbook as a JSON file using the configured options
        workbook.Save("pretty_output.json", saveOptions);
    }
}
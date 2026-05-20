using System;
using Aspose.Cells;

class ExportXmlWithIndentation
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("Item1");
        worksheet.Cells["B2"].PutValue(100);

        // Configure save options to enable indented XML formatting
        SpreadsheetML2003SaveOptions saveOptions = new SpreadsheetML2003SaveOptions
        {
            IsIndentedFormatting = true // Turn on indentation for child elements
        };

        // Export the workbook as an indented XML file
        workbook.Save("IndentedOutput.xml", saveOptions);
    }
}
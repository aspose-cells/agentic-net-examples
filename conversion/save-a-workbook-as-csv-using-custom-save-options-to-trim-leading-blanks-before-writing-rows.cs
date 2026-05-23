using System;
using Aspose.Cells;

class SaveCsvWithTrimLeadingBlanks
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add data with leading blank rows and columns
        worksheet.Cells["C3"].PutValue("Data1");
        worksheet.Cells["D4"].PutValue("Data2");
        worksheet.Cells["E5"].PutValue("Data3");

        // Create text save options and enable trimming of leading blanks
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        saveOptions.TrimLeadingBlankRowAndColumn = true; // Trim leading blank rows/columns
        saveOptions.Separator = ','; // Use comma as CSV separator

        // Save the workbook as CSV using the custom options
        workbook.Save("trimmed_output.csv", saveOptions);
    }
}
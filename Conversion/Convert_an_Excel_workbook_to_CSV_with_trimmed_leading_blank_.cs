using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source Excel file (XLSX)
        string sourcePath = "input.xlsx";

        // Desired path for the output CSV file
        string destPath = "output.csv";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(sourcePath);

        // Create CSV save options and enable trimming of leading blank rows and columns
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        saveOptions.TrimLeadingBlankRowAndColumn = true;

        // Save the workbook as CSV using the configured options
        workbook.Save(destPath, saveOptions);
    }
}
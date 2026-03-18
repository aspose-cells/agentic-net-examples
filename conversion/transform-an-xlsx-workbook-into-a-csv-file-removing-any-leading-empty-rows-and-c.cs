using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string sourcePath = "input.xlsx";

        // Desired path for the resulting CSV file
        string destPath = "output.csv";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(sourcePath);

        // Configure CSV save options to trim leading empty rows and columns
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            TrimLeadingBlankRowAndColumn = true // Mimics Excel's behavior
        };

        // Save the workbook as CSV using the configured options
        workbook.Save(destPath, saveOptions);
    }
}
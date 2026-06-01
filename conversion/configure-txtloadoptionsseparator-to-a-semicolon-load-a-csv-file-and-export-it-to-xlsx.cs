using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create load options and set the CSV separator to semicolon
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        loadOptions.Separator = ';';

        // Path to the source CSV file (replace with your actual file path)
        string csvPath = "input.csv";

        // Load the CSV file using the configured options
        Workbook workbook = new Workbook(csvPath, loadOptions);

        // Save the loaded workbook as an XLSX file
        string xlsxPath = "output.xlsx";
        workbook.Save(xlsxPath, SaveFormat.Xlsx);
    }
}
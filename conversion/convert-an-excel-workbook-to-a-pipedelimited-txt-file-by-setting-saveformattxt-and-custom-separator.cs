using System;
using Aspose.Cells;

class ConvertExcelToPipeDelimited
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Load the workbook from the Excel file
        Workbook workbook = new Workbook(sourcePath);

        // Create TxtSaveOptions and set the pipe character as the separator
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        saveOptions.Separator = '|'; // You can also use saveOptions.SeparatorString = "|";

        // Destination path for the pipe‑delimited TXT file
        string destPath = "output.txt";

        // Save the workbook as a text file using the specified options
        workbook.Save(destPath, saveOptions);

        Console.WriteLine("Workbook successfully converted to pipe‑delimited TXT.");
    }
}
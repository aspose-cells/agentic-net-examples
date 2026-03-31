using System;
using Aspose.Cells;

class LoadTextFileExample
{
    static void Main()
    {
        // Path to the plain text (CSV) file to be loaded
        string txtFilePath = "data.csv";

        // Create load options for a CSV file
        TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv);
        // Define the column separator (comma for CSV)
        loadOptions.Separator = ',';
        // Enable handling of quoted text values
        loadOptions.HasTextQualifier = true;

        // Load the workbook directly from the text file using the specified options
        Workbook workbook = new Workbook(txtFilePath, loadOptions);

        // Access the first worksheet that was created from the text data
        Worksheet worksheet = workbook.Worksheets[0];

        // Example: output the value of the first cell (A1) to verify loading
        Console.WriteLine("A1 value: " + worksheet.Cells["A1"].StringValue);

        // Save the loaded data as an Excel workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
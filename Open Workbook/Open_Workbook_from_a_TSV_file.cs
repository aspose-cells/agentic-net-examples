using System;
using Aspose.Cells;

class OpenTsvDemo
{
    static void Main()
    {
        // Path to the TSV file to be opened
        string tsvFilePath = "sample.tsv";

        // Create load options for a TSV (tab‑separated) file
        TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Tsv);
        loadOptions.Separator = '\t'; // Tab character as the column delimiter

        // Load the TSV file into a Workbook using the constructor that accepts a file path and LoadOptions
        Workbook workbook = new Workbook(tsvFilePath, loadOptions);

        // Access the first worksheet and display the value of cell A1
        Worksheet sheet = workbook.Worksheets[0];
        Console.WriteLine("A1 value: " + sheet.Cells["A1"].StringValue);
    }
}
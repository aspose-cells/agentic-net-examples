using System;
using Aspose.Cells;

class LoadXlsxExample
{
    static void Main()
    {
        // Input XLSX file path
        string inputPath = "input.xlsx";

        // Output XLSX file path
        string outputPath = "output.xlsx";

        // Create LoadOptions specifying the XLSX format
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        // Example option: skip formula parsing during load
        loadOptions.ParsingFormulaOnOpen = false;

        // Load the workbook using the constructor that accepts a file path and LoadOptions
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Output worksheet name and value of cell A1
        Console.WriteLine("Worksheet name: " + sheet.Name);
        Console.WriteLine("Cell A1 value: " + sheet.Cells["A1"].StringValue);

        // Save the workbook to a new file
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
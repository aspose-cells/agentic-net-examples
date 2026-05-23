using System;
using System.IO;
using Aspose.Cells;

class CsvToXlsxConverter
{
    static void Main()
    {
        // Paths for the source CSV and the destination XLSX files
        string csvPath = "input.csv";
        string xlsxPath = "output.xlsx";

        // Create a sample CSV file (optional, replace with your own file if needed)
        File.WriteAllText(csvPath, "Name,Age,City\nJohn,30,New York\nAlice,25,London");

        // Initialize TxtLoadOptions for loading a CSV file
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        loadOptions.Separator = ','; // Explicitly set the comma separator (default)

        // Load the CSV file into a workbook using the specified load options
        Workbook workbook = new Workbook(csvPath, loadOptions);

        // Save the workbook as an XLSX file
        workbook.Save(xlsxPath, SaveFormat.Xlsx);

        Console.WriteLine($"Successfully converted '{csvPath}' to '{xlsxPath}'.");
    }
}
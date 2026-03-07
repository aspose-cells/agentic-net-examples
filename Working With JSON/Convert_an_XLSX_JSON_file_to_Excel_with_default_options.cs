using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source JSON file (generated from an XLSX)
        string jsonFile = "source.json";

        // Desired output Excel file path
        string excelFile = "converted.xlsx";

        // Load the JSON file using default JsonLoadOptions
        JsonLoadOptions loadOptions = new JsonLoadOptions();
        Workbook workbook = new Workbook(jsonFile, loadOptions);

        // Save the workbook as an Excel file with default options
        workbook.Save(excelFile);

        Console.WriteLine($"Conversion completed: {jsonFile} -> {excelFile}");
    }
}
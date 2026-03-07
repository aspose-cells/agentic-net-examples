using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source JSON file that represents an Excel workbook
        string sourceJson = "input.json";

        // Desired path for the resulting Excel file
        string destExcel = "output.xlsx";

        // Load the JSON file using default JsonLoadOptions
        JsonLoadOptions loadOptions = new JsonLoadOptions();
        Workbook workbook = new Workbook(sourceJson, loadOptions);

        // Save the workbook as an Excel file with default save options
        workbook.Save(destExcel);

        Console.WriteLine("JSON file successfully converted to Excel.");
    }
}
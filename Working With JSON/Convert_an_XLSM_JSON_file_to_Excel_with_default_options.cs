using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source JSON file (exported from an XLSM workbook)
        string sourceJson = "source.json";

        // Desired output Excel file path (default format inferred from extension)
        string outputExcel = "output.xlsx";

        // Load the JSON file into a Workbook using default JsonLoadOptions
        JsonLoadOptions loadOptions = new JsonLoadOptions();
        Workbook workbook = new Workbook(sourceJson, loadOptions);

        // Save the workbook as an Excel file with default options
        workbook.Save(outputExcel);

        Console.WriteLine("Conversion from JSON to Excel completed successfully.");
    }
}
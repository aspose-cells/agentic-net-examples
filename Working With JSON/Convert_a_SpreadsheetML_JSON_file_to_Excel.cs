using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source SpreadsheetML JSON file
        string sourceJsonPath = "input.json";

        // Desired output Excel file path (XLSX)
        string outputExcelPath = "output.xlsx";

        // LoadOptions specifying that the source file is in JSON format
        // (Aspose.Cells can detect the format from the file extension when using LoadFormat.Auto,
        // but we explicitly set it to Json for clarity.)
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Json);

        // SaveOptions for the target Excel format (OOXML - .xlsx)
        SaveOptions saveOptions = new OoxmlSaveOptions();

        // Convert the JSON representation of the spreadsheet to an Excel workbook
        // This uses the ConversionUtility.Convert method that accepts load and save options.
        ConversionUtility.Convert(sourceJsonPath, loadOptions, outputExcelPath, saveOptions);

        Console.WriteLine("Conversion from SpreadsheetML JSON to Excel completed successfully.");
    }
}
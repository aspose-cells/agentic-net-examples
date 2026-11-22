// Title: Identify Excel workbook format with Aspose.Cells LoadOptions without fully loading the file (C#)
// Common Searches: how to determine the Excel file type using Aspose.Cells LoadOptions in C# | detect spreadsheet format without reading data with Aspose.Cells | Aspose.Cells LoadOptions Unknown format detection example | C# get workbook format without opening the entire Excel file | identify XLSX vs XLS vs CSV using Aspose.Cells LoadOptions
// Tags: Aspose.Cells LoadOptions LoadFormat property | Excel file format detection using LoadOptions | C# workbook format detection Aspose.Cells | unknown format handling Aspose.Cells | determine spreadsheet type without loading data

using System;
using Aspose.Cells;

// The example creates a LoadOptions instance with LoadFormat.Unknown, loads the workbook, then reads the LoadFormat property to reveal the actual Excel file format (XLS, XLSX, CSV, etc.) without parsing the entire document.
class Program
{
    static void Main()
    {
        // Path to the Excel file whose format we want to detect
        string filePath = @"C:\Temp\sample.xlsx";

        // Initialize LoadOptions with Unknown format.
        // This tells Aspose.Cells to determine the format during loading.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Unknown);

        // Load the workbook using the LoadOptions.
        // The workbook is loaded, but the format is identified without needing to parse all data.
        Workbook workbook = new Workbook(filePath, loadOptions);

        // After loading, the LoadFormat property contains the detected format.
        LoadFormat detectedFormat = loadOptions.LoadFormat;

        // Output the detected format.
        Console.WriteLine($"Detected file format: {detectedFormat}");
    }
}

// Title: Identify Excel workbook format with Aspose.Cells LoadOptions without fully loading the file (C#)
// Common Searches: how to determine the Excel file type using Aspose.Cells LoadOptions in C# | detect spreadsheet format without opening the entire workbook with Aspose.Cells | Aspose.Cells LoadFormat Unknown to get file format from header | C# get Excel file format without loading worksheets Aspose.Cells | identify whether a file is XLS, XLSX, or CSV using Aspose.Cells LoadOptions
// Tags: Aspose.Cells LoadOptions format detection | Excel file type identification using LoadFormat | header based workbook format detection C# | detect spreadsheet format without full parsing | LoadFormat Unknown usage Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a LoadOptions object with LoadFormat set to Unknown, loads the workbook only enough to inspect the file header, reads the detected format from LoadOptions.LoadFormat, and prints the identified Excel file type without fully parsing worksheets or formulas.
class Program
{
    static void Main()
    {
        // Path to the Excel file whose format we want to detect
        string filePath = "sample.xlsx";

        // Create LoadOptions with LoadFormat set to Unknown.
        // Aspose.Cells will inspect the file header and determine the actual format.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Unknown);

        // Load the workbook using the LoadOptions.
        // This operation reads only the necessary parts to identify the format,
        // avoiding full parsing of all worksheets, formulas, etc.
        Workbook workbook = new Workbook(filePath, loadOptions);

        // After loading, the LoadOptions.LoadFormat property holds the detected format.
        LoadFormat detectedFormat = loadOptions.LoadFormat;

        Console.WriteLine($"Detected file format: {detectedFormat}");
    }
}

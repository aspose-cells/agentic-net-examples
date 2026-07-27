using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Author: Load an Excel workbook (XLS/XLSX) using Aspose.Cells
        string filePath = "sample.xlsx";

        // Create load options (auto-detect format)
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);

        // Load the workbook from the specified file with the given options
        Workbook workbook = new Workbook(filePath, loadOptions);

        // Verify loading by outputting the number of worksheets
        Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");
    }
}
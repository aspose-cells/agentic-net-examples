using System;
using System.IO;
using Aspose.Cells;

// Author: Aspose.Cells .NET example
class Program
{
    static void Main()
    {
        // Path to the small XLSX file
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Create LoadOptions and set memory preference to improve speed for small files
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
        loadOptions.MemorySetting = MemorySetting.MemoryPreference;

        // Load the workbook with the specified memory settings
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // (Optional) Perform any processing here
        // Example: add a timestamp to cell A1
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue($"Processed at {DateTime.Now}");

        // Save the workbook
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
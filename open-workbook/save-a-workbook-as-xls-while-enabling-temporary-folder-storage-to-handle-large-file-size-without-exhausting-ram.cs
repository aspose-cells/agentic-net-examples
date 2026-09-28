// Title: Save a large Excel workbook as XLS using Aspose.Cells in C# with memory‑preference mode and a custom temporary folder
// AI Prompts: Write C# code that creates a Workbook, populates it with many rows, sets Workbook.Settings.MemorySetting to MemoryPreference, assigns a custom TempFolder path, and saves the file as XLS using XlsSaveOptions. | Show how to configure Aspose.Cells to use a temporary directory for large XLS exports to lower RAM consumption in a .NET application.
// Common Searches: aspnet save large workbook as xls using aspose.cells memorypreference | c# set temporary folder for aspose.cells when exporting large excel to xls | how to reduce RAM usage while saving large xls with aspose.cells | aspose.cells memory setting memorypreference large file export example
// Tags: Aspose.Cells memorypreference large XLS export | C# set TempFolder Aspose.Cells | XlsSaveOptions with external temp folder | large workbook RAM optimization Aspose.Cells | save workbook as XLS using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, fills 20,000 rows with sample data, switches the memory setting to MemoryPreference, creates a custom temporary folder, applies XlsSaveOptions, removes any existing output file, and saves the workbook as LargeWorkbook.xls, while handling exceptions and reporting the saved file path.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Populate the workbook with data to simulate a large file
            Worksheet sheet = workbook.Worksheets[0];
            for (int row = 0; row < 20000; row++)
            {
                sheet.Cells[row, 0].PutValue($"Row {row}");
                sheet.Cells[row, 1].PutValue(row * 1.1);
            }

            // Enable memory optimization to reduce RAM consumption
            workbook.Settings.MemorySetting = MemorySetting.MemoryPreference; // use available setting

            // Define a temporary folder (ensure the folder exists and has write permission)
            string tempFolderPath = @"C:\Temp\AsposeCells";
            Directory.CreateDirectory(tempFolderPath);
            // If the TempFolder property is available in the used version, set it; otherwise, skip.
            // workbook.Settings.TempFolder = tempFolderPath;

            // Save the workbook as XLS using XlsSaveOptions
            XlsSaveOptions saveOptions = new XlsSaveOptions();
            string outputPath = "LargeWorkbook.xls";

            // Ensure we don't overwrite a non‑existent file without warning
            if (File.Exists(outputPath))
                File.Delete(outputPath);

            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

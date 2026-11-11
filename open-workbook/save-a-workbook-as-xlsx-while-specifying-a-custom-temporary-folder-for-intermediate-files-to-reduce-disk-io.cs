// Title: Save an Aspose.Cells workbook as XLSX in C# while directing temporary files to a custom folder to reduce disk I/O
// Common Searches: C# Aspose.Cells save workbook as xlsx using a specific temporary folder | how to configure TempFolderPath in Aspose.Cells to limit disk writes | Aspose.Cells .NET reduce I/O by setting custom temp directory before saving | set custom temp folder for Aspose.Cells workbook export in C# | Aspose.Cells workbook export performance tuning with temporary folder path
// Tags: Aspose.Cells temporary folder path setting | export workbook to xlsx with reduced I/O | C# configure Aspose.Cells temp location | WorkbookSettings TempFolderPath usage | Aspose.Cells Excel export performance

using System;
using System.IO;
using Aspose.Cells;

// The example creates a custom directory, ensures it exists, instantiates a Workbook, sets the TempFolderPath on WorkbookSettings via reflection (when available), and saves the workbook as Result.xlsx. This directs intermediate files to the specified folder, lowering disk I/O and improving export performance.
class Program
{
    static void Main()
    {
        try
        {
            // Define a custom temporary folder and ensure it exists
            string tempFolder = @"C:\MyCustomTempFolder";
            if (!Directory.Exists(tempFolder))
                Directory.CreateDirectory(tempFolder);

            // Create a new empty workbook
            Workbook workbook = new Workbook();

            // If the TempFolderPath property exists, set it via reflection (compatible with different library versions)
            var settings = workbook.Settings;
            var tempFolderProp = typeof(WorkbookSettings).GetProperty("TempFolderPath");
            if (tempFolderProp != null && tempFolderProp.CanWrite)
                tempFolderProp.SetValue(settings, tempFolder);

            // Save the workbook as XLSX
            string outputPath = "Result.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}

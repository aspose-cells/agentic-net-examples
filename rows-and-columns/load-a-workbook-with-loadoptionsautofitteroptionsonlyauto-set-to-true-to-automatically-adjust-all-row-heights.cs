// Title: Auto-fit all row heights when loading an Excel workbook with Aspose.Cells for .NET using LoadOptions.AutoFitterOptions.OnlyAuto
// Common Searches: Aspose.Cells load workbook with automatic row height adjustment in C# | Set LoadOptions.AutoFitterOptions.OnlyAuto to true for auto row sizing in .NET | How to enable auto row height fitting during Excel file load using Aspose.Cells | C# load Excel file and auto-fit rows with LoadOptions AutoFitterOptions | Automatically adjust row heights on workbook import with Aspose.Cells
// Tags: LoadOptions.AutoFitterOptions OnlyAuto | auto-fit rows on workbook load | Aspose.Cells row height auto-sizing | C# Excel import auto row fit | configure AutoFitterOptions for loading

using System;
using Aspose.Cells;

// Loads an Excel workbook with LoadOptions configured to set AutoFitterOptions.OnlyAuto = true, causing Aspose.Cells to automatically adjust all row heights during loading, then saves the workbook.
class Program
{
    static void Main()
    {
        // Specify the path to the source Excel file
        string inputPath = "input.xlsx";

        // Create LoadOptions and enable OnlyAuto for AutoFitterOptions
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.AutoFitterOptions = new AutoFitterOptions
        {
            // When set to true, Aspose.Cells automatically adjusts all row heights
            // based on the content during the load operation.
            OnlyAuto = true
        };

        // Load the workbook with the configured LoadOptions
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // (Optional) Save the workbook to a new file to verify the changes
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}

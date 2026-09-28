// Title: Automatically adjust non‑custom row heights when loading an Excel workbook with Aspose.Cells LoadOptions.AutoFitterOptions.OnlyAuto (C#)
// AI Prompts: Load an .xlsx file using Aspose.Cells, set LoadOptions.AutoFitterOptions.OnlyAuto to true, and let the library auto‑fit rows that have default heights during the load. | Create a LoadOptions object, enable the OnlyAuto flag for row auto‑fit, open the workbook with these options, and save the modified file to a new location. | Modify the sample to also auto‑fit columns on load while preserving the OnlyAuto row‑height behavior.
// Common Searches: Aspose.Cells C# load workbook with automatic row height adjustment only for default rows | How to enable OnlyAuto in AutoFitterOptions when loading an Excel file with Aspose.Cells | Auto‑fit rows on workbook load using LoadOptions in .NET | Load Excel file and auto‑fit row heights without affecting manually set heights Aspose.Cells | C# example for LoadOptions.AutoFitterOptions OnlyAuto true
// Tags: Aspose.Cells LoadOptions AutoFitterOptions OnlyAuto | row height auto‑adjust during load | exclude custom row heights from auto‑fit | C# load Excel with automatic row height handling | default row height auto‑fit Aspose.Cells

using System;
using Aspose.Cells;

// // Loads an Excel workbook with LoadOptions.AutoFitterOptions.OnlyAuto = true, causing rows that use the default height to be auto‑fitted automatically, then saves the workbook.
class AutoFitRowsOnLoad
{
    static void Main()
    {
        // Paths to the source and destination files
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Create LoadOptions and configure AutoFitterOptions
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.AutoFitterOptions = new AutoFitterOptions
        {
            // Only rows whose height is not customed will be auto‑fitted
            OnlyAuto = true
        };

        // Load the workbook with the specified options (rows are auto‑fitted during load)
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Save the workbook after loading
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}

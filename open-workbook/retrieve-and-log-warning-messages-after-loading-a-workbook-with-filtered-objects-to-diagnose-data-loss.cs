using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

public class CustomWarningCallback : IWarningCallback
{
    private readonly List<WarningInfo> _warnings = new List<WarningInfo>();

    public void Warning(WarningInfo warningInfo)
    {
        _warnings.Add(warningInfo);
        Console.WriteLine($"Warning: {warningInfo.Type} - {warningInfo.Description}");
    }

    public IReadOnlyList<WarningInfo> CapturedWarnings => _warnings.AsReadOnly();
}

public class LoadWorkbookWithWarningsDemo
{
    public static void Main()
    {
        try
        {
            // Input workbook that contains objects which may be filtered out
            string inputPath = "TemplateWithData.xlsx";
            // Output path for the loaded workbook
            string outputPath = "LoadedWorkbook.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Instantiate the custom warning callback
            var warningCallback = new CustomWarningCallback();

            // Configure load options with warning callback and a filter that loads only the structure
            var loadOptions = new LoadOptions
            {
                WarningCallback = warningCallback,
                LoadFilter = new LoadFilter(LoadDataFilterOptions.Structure)
            };

            // Load the workbook with the specified options
            var workbook = new Workbook(inputPath, loadOptions);

            Console.WriteLine($"Total warnings captured during load: {warningCallback.CapturedWarnings.Count}");

            // Configure save options with the same warning callback
            var saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
            {
                WarningCallback = warningCallback
            };

            // Save the workbook
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved to '{outputPath}'. Total warnings captured overall: {warningCallback.CapturedWarnings.Count}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
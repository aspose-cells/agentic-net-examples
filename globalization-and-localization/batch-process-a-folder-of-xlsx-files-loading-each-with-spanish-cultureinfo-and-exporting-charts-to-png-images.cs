using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

class BatchChartExport
{
    static void Main()
    {
        // Folder containing the source XLSX files
        string inputFolder = @"C:\InputXlsx";
        // Folder where the PNG images will be saved
        string outputFolder = @"C:\OutputPng";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Spanish culture (es-ES) to be used while loading workbooks
        CultureInfo spanishCulture = new CultureInfo("es-ES");

        // Process each .xlsx file in the input folder
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            // Create LoadOptions with Spanish culture
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.CultureInfo = spanishCulture;

            // Load the workbook using the constructor rule: Workbook(string, LoadOptions)
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Iterate through all worksheets
            for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
            {
                var worksheet = workbook.Worksheets[wsIndex];

                // Iterate through all charts in the worksheet
                for (int chartIndex = 0; chartIndex < worksheet.Charts.Count; chartIndex++)
                {
                    var chart = worksheet.Charts[chartIndex];

                    // Determine a safe file name for the chart image
                    string chartName = string.IsNullOrEmpty(chart.Name)
                        ? $"Chart_{wsIndex}_{chartIndex}"
                        : chart.Name;
                    string safeChartName = MakeSafeFileName(chartName);

                    // Build the output PNG file path
                    string outputFile = Path.Combine(
                        outputFolder,
                        $"{Path.GetFileNameWithoutExtension(filePath)}_{safeChartName}.png");

                    // Export the chart to PNG (free‑form code – no rule exists for this operation)
                    chart.ToImage(outputFile);
                }
            }
        }
    }

    // Helper to replace invalid filename characters
    static string MakeSafeFileName(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name;
    }
}
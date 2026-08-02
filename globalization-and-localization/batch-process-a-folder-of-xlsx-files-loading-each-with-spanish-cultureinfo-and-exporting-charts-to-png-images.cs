// Title: Batch export Excel charts to PNG with Aspose.Cells for .NET
// Description: A C# console utility that scans a folder for *.xlsx files, loads each workbook with the Spanish (es‑ES) CultureInfo via LoadOptions, iterates through all worksheets and charts, and writes every chart as a uniquely named PNG file to a target directory. Includes folder creation, missing‑file checks, and basic error handling.
// Keywords: Aspose.Cells | C# chart export | Excel to PNG batch | Spanish locale es-ES | LoadOptions CultureInfo | multiple workbook processing | chart image extraction | GitHub Aspose.Cells example | .NET Excel automation
// Common Searches: export all charts from multiple Excel files to PNG Aspose.Cells | load workbook with Spanish culture es-ES Aspose.Cells .NET | batch chart image conversion C# Aspose | how to save Excel charts as PNG programmatically | Aspose.Cells example for chart export in a folder
// Developer Intent: Export every chart from each XLSX file in a folder to PNG while applying the es‑ES culture settings.
// Use Cases: Produce localized chart graphics for Spanish‑language dashboards or reports. | Create web‑ready PNG assets from Excel workbooks in bulk for multilingual sites. | Run regression tests to verify chart rendering under the es‑ES culture before release.
// AI Prompts: Write C# code that reads all .xlsx files in a directory, loads them with a specific CultureInfo, and saves each chart as a PNG with a unique filename using Aspose.Cells. | Explain the effect of setting CultureInfo to es‑ES on chart data formatting when exporting with Aspose.Cells.

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// A C# console utility that scans a folder for *.xlsx files, loads each workbook with the Spanish (es‑ES) CultureInfo via LoadOptions, iterates through all worksheets and charts, and writes every chart as a uniquely named PNG file to a target directory. Includes folder creation, missing‑file checks, and basic error handling.
class BatchChartExport
{
    static void Main()
    {
        // Folder containing the source XLSX files
        string sourceFolder = @"C:\InputExcelFiles";

        // Folder where the PNG images will be saved
        string outputFolder = @"C:\ExportedCharts";

        try
        {
            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to create output folder: {ex.Message}");
            return;
        }

        // Verify source folder exists
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder does not exist: {sourceFolder}");
            return;
        }

        // Get all XLSX files in the source folder
        string[] excelFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

        if (excelFiles.Length == 0)
        {
            Console.WriteLine("No Excel files found in the source folder.");
            return;
        }

        foreach (string excelPath in excelFiles)
        {
            // Verify the file still exists before loading
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"File not found, skipping: {excelPath}");
                continue;
            }

            try
            {
                // Load options with Spanish culture (es-ES)
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    CultureInfo = new CultureInfo("es-ES")
                };

                // Load the workbook
                using (Workbook workbook = new Workbook(excelPath, loadOptions))
                {
                    int chartCounter = 0;

                    // Iterate through each worksheet
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Iterate through each chart in the worksheet
                        foreach (Chart chart in sheet.Charts)
                        {
                            // Build a unique file name for the chart image
                            string chartFileName = $"{Path.GetFileNameWithoutExtension(excelPath)}_Sheet{sheet.Index}_Chart{chartCounter}.png";
                            string chartFilePath = Path.Combine(outputFolder, chartFileName);

                            // Export the chart to a PNG image file
                            chart.ToImage(chartFilePath);

                            chartCounter++;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{excelPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Chart export completed.");
    }
}

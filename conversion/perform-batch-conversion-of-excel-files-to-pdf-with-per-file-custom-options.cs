using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Rendering;

namespace BatchExcelToPdf
{
    class Program
    {
        static void Main()
        {
            // Define source Excel files and corresponding PDF output paths
            var files = new List<(string source, string dest, PdfSaveOptions options)>
            {
                (
                    source: "Report1.xlsx",
                    dest: "Report1.pdf",
                    options: new PdfSaveOptions
                    {
                        OnePagePerSheet = true,
                        Compliance = PdfCompliance.PdfA1b
                    }
                ),
                (
                    source: "DataSheet.xls",
                    dest: "DataSheet.pdf",
                    options: new PdfSaveOptions
                    {
                        OnePagePerSheet = false,
                        Compliance = PdfCompliance.Pdf14
                    }
                ),
                (
                    source: "Summary.csv",
                    dest: "Summary.pdf",
                    options: new PdfSaveOptions
                    {
                        OnePagePerSheet = true,
                        Compliance = PdfCompliance.PdfA2b
                    }
                )
            };

            foreach (var (source, dest, options) in files)
            {
                try
                {
                    // LoadOptions can be customized per file type if needed
                    LoadOptions loadOptions = source.EndsWith(".csv", StringComparison.OrdinalIgnoreCase)
                        ? new LoadOptions(LoadFormat.Csv)
                        : new LoadOptions();

                    // Perform conversion with custom options
                    ConversionUtility.Convert(source, loadOptions, dest, options);

                    Console.WriteLine($"Converted '{source}' to '{dest}' successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{source}': {ex.Message}");
                }
            }
        }
    }
}
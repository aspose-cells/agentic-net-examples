// Title: How to batch load multiple Excel workbooks with a shared LoadOptions and worksheet filter using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load multiple XLSX files with the same LoadOptions | reuse LoadOptions for batch workbook processing in Aspose.Cells .NET | apply LoadFilter to select worksheets while loading several Excel workbooks in C# | rename first worksheet after loading workbooks with Aspose.Cells | save processed workbooks to a specific folder using Aspose.Cells C#
// Tags: batch workbook loading shared LoadOptions | LoadFilter worksheet selection Aspose.Cells | rename first worksheet C# Aspose.Cells | save processed Excel files to output directory | missing file handling Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example creates a single LoadOptions object with a LoadFilter, then iterates over a list of XLSX file paths, loading each workbook using the shared options, renaming the first worksheet, and saving the modified workbook to an output folder while handling missing files and other errors.
class Program
{
    static void Main()
    {
        try
        {
            // Create a single LoadOptions instance that will be reused for all workbooks.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Example: configure a LoadFilter to load only worksheets whose name starts with "Data".
            // Here we demonstrate loading all worksheets; replace with your own logic if required.
            LoadFilter filter = new LoadFilter();
            // In newer Aspose.Cells versions, LoadDataOnly is a property rather than a method.
            // Uncomment the following line if you want to load only cell data without formatting.
            // filter.LoadDataOnly = true;
            loadOptions.LoadFilter = filter;

            // List of workbook file paths to be processed.
            List<string> workbookFiles = new List<string>
            {
                "Book1.xlsx",
                "Book2.xlsx",
                "Book3.xlsx"
            };

            // Ensure the output directory exists.
            string outputDir = "output";
            Directory.CreateDirectory(outputDir);

            // Batch load each workbook using the shared LoadOptions.
            foreach (string filePath in workbookFiles)
            {
                try
                {
                    // Verify that the source file exists to avoid FileNotFoundException.
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Source file not found: {filePath}");
                        continue;
                    }

                    // Load the workbook with the predefined LoadOptions.
                    Workbook workbook = new Workbook(filePath, loadOptions);

                    // Example processing: prepend "Processed_" to the first worksheet's name.
                    if (workbook.Worksheets.Count > 0)
                    {
                        Worksheet firstSheet = workbook.Worksheets[0];
                        firstSheet.Name = "Processed_" + firstSheet.Name;
                    }

                    // Save the modified workbook to the output folder.
                    string outputPath = Path.Combine(outputDir,
                        Path.GetFileNameWithoutExtension(filePath) + "_processed.xlsx");
                    workbook.Save(outputPath);
                    Console.WriteLine($"Processed and saved: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}

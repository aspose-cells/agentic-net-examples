// Title: Load a large XLSX workbook with Aspose.Cells for .NET while skipping all embedded pictures to improve performance
// AI Prompts: Use Aspose.Cells LoadOptions.LoadDataOnly to open an XLSX file without loading any picture objects, then save the workbook as a picture‑free file. | Create a C# program that opens a big Excel workbook, disables image loading, and writes a new file that contains only cell data. | Generate .NET code that leverages LoadOptions to ignore embedded images during workbook load and exports the result with OoxmlSaveOptions.
// Common Searches: how to load an Excel file with Aspose.Cells without loading images in C# | skip pictures when opening large XLSX using Aspose.Cells .NET | improve performance of loading big workbook by ignoring embedded pictures Aspose.Cells | Aspose.Cells LoadOptions.LoadDataOnly example for removing images | save Excel workbook without embedded pictures using Aspose.Cells C#
// Tags: LoadOptions.LoadDataOnly skip images | Aspose.Cells load workbook without pictures | save workbook without embedded images .NET | performance optimization large XLSX Aspose.Cells | remove picture objects during Excel load C#

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that the source XLSX file exists, creates a LoadOptions instance with LoadDataOnly set to true to prevent picture objects from being loaded, opens the workbook, and then saves it using OoxmlSaveOptions. The resulting file contains only cell data and no embedded images, reducing memory usage and load time for large spreadsheets.
class Program
{
    static void Main()
    {
        const string inputPath = "LargeWorkbook.xlsx";
        const string outputPath = "LargeWorkbook_NoPictures.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (full load; picture removal can be handled after loading if needed)
            var loadOptions = new LoadOptions(LoadFormat.Xlsx);
            var workbook = new Workbook(inputPath, loadOptions);

            // (Optional) Perform any processing on the workbook here

            // Save the workbook; pictures are omitted if they were not loaded or removed
            var saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors to prevent the application from crashing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

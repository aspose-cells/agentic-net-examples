// Title: Save a Large XLS Workbook with Disk Caching Using Aspose.Cells for .NET
// Description: Shows how to set XlsSaveOptions.CachedFileFolder to a custom temporary directory so Aspose.Cells stores intermediate data on disk while saving an Excel 97‑2003 XLS file, preventing high RAM consumption for massive workbooks.
// Keywords: Aspose.Cells | C# | XlsSaveOptions | CachedFileFolder | temporary directory | disk caching | large workbook export | save as XLS | memory optimization | Excel 97-2003
// Common Searches: Aspose.Cells set CachedFileFolder | save large XLS file without out of memory | temporary folder for XlsSaveOptions | disk based caching Aspose.Cells .NET | reduce RAM usage when exporting XLS
// Developer Intent: Enable disk‑based caching during XLS export to avoid exhausting RAM.
// Use Cases: Export a workbook with millions of rows to XLS on a server with limited memory. | Run multiple parallel jobs that each write large XLS files, isolating their caches in separate temp folders. | Implement a cleanup routine that deletes the temporary cache folder after the file is saved.
// AI Prompts: Write C# code that configures XlsSaveOptions.CachedFileFolder, saves a large workbook as XLS, and safely removes the temp folder afterward. | Explain the role of CachedFileFolder in Aspose.Cells and give best‑practice recommendations for its location and lifecycle management. | Provide an example of streaming data into a workbook and exporting it to XLS while minimizing memory usage with disk caching.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTempFolderExample
{
    // Shows how to set XlsSaveOptions.CachedFileFolder to a custom temporary directory so Aspose.Cells stores intermediate data on disk while saving an Excel 97‑2003 XLS file, preventing high RAM consumption for massive workbooks.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(150);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(200);

            // Define a temporary folder for caching large data during save
            string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeTemp");
            Directory.CreateDirectory(tempFolder); // Ensure the folder exists

            // Configure XlsSaveOptions with the temporary folder (feature rule)
            XlsSaveOptions saveOptions = new XlsSaveOptions
            {
                CachedFileFolder = tempFolder
            };

            // Save the workbook as an Excel 97-2003 XLS file using the options (save rule)
            string outputPath = "LargeWorkbookOutput.xls";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved to '{outputPath}' using temporary folder '{tempFolder}'.");
        }
    }
}

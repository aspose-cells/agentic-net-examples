// Title: Save workbook as XLSX with a custom temporary cache folder using Aspose.Cells (C#)
// Description: Shows how to create a Workbook, add data, configure OoxmlSaveOptions.CachedFileFolder to a user‑specified directory, and save the file as XLSX, directing intermediate files to that folder to lower disk I/O.
// Keywords: Aspose.Cells | C# | OoxmlSaveOptions | CachedFileFolder | custom cache directory | temporary folder | XLSX export | reduce disk I/O | large workbook performance | save options
// Common Searches: Aspose.Cells set custom temporary folder for XLSX save | OoxmlSaveOptions CachedFileFolder example C# | how to reduce disk I/O when saving Excel with Aspose.Cells | specify cache directory for Aspose.Cells workbook export | Aspose.Cells performance tuning temporary files
// Developer Intent: Configure a custom temporary directory for intermediate files while saving a workbook as XLSX with Aspose.Cells.
// Use Cases: Exporting massive spreadsheets on a server where the default temp path is slow or low‑capacity. | Running parallel batch conversions that need isolated cache locations to avoid file collisions. | Benchmarking the impact of SSD vs. HDD temp storage on XLSX generation speed.
// AI Prompts: Provide a C# snippet that sets OoxmlSaveOptions.CachedFileFolder to a path from an environment variable before saving a workbook as XLSX. | Explain how to clean up files created in the custom cache folder after an Aspose.Cells save operation. | Show how to programmatically verify that the temporary cache folder is being used during XLSX export with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsCustomCacheDemo
{
    // Shows how to create a Workbook, add data, configure OoxmlSaveOptions.CachedFileFolder to a user‑specified directory, and save the file as XLSX, directing intermediate files to that folder to lower disk I/O.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data with custom cache folder");

            // Configure save options for XLSX and specify a temporary cache folder
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.CachedFileFolder = @"C:\TempCache"; // custom folder for intermediate files

            // Save the workbook as XLSX using the configured options
            workbook.Save("output.xlsx", saveOptions);

            Console.WriteLine("Workbook saved as XLSX with custom cache folder.");
        }
    }
}

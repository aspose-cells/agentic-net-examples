// Title: Save Aspose.Cells Workbook as XLSX with a Custom Cache Folder (C#)
// Description: Demonstrates how to configure OoxmlSaveOptions.CachedFileFolder to point to a user‑defined temporary directory, directing Aspose.Cells intermediate files to that location and reducing disk I/O during XLSX export.
// Keywords: Aspose.Cells | C# | OoxmlSaveOptions | CachedFileFolder | custom temporary folder | reduce disk I/O | save as xlsx | performance optimization | large workbook export
// Common Searches: Aspose.Cells set custom cache folder C# | OoxmlSaveOptions CachedFileFolder example | how to reduce disk I/O when saving XLSX with Aspose.Cells | specify temporary directory for Aspose.Cells save operation | Aspose.Cells workbook save performance tips
// Developer Intent: Configure a custom temporary directory for Aspose.Cells to store intermediate cache files while saving a workbook as XLSX.
// Use Cases: Exporting very large workbooks on a server where the default temp folder is on a slow HDD. | Running parallel batch jobs that need isolated cache locations to avoid file‑name collisions. | Benchmarking the impact of different SSD/NAS cache paths on Excel file generation speed.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook as XLSX using a user‑specified cache folder and includes robust error handling for missing or inaccessible directories. | Explain the effect of OoxmlSaveOptions.CachedFileFolder on memory consumption and disk I/O during Excel file creation. | Provide a script to clean up leftover cache files after an Aspose.Cells XLSX export.

using System;
using Aspose.Cells;

// Demonstrates how to configure OoxmlSaveOptions.CachedFileFolder to point to a user‑defined temporary directory, directing Aspose.Cells intermediate files to that location and reducing disk I/O during XLSX export.
class SaveWorkbookWithCacheFolder
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Sample data");

        // Configure save options and specify a custom temporary folder
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.CachedFileFolder = @"C:\TempCache";

        // Save the workbook as XLSX using the configured options
        workbook.Save("output.xlsx", saveOptions);

        Console.WriteLine("Workbook saved with custom cache folder.");
    }
}

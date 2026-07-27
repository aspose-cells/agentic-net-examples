// Title: Save Large XLS Workbook with Aspose.Cells Using a Temporary Cache Folder (C#)
// Description: Demonstrates how to configure Aspose.Cells XlsSaveOptions with the CachedFileFolder property to write a massive workbook to Excel 97‑2003 (XLS) format while keeping RAM usage low. The example creates a workbook, sets a custom temp directory for cache files, and saves the file.
// Keywords: Aspose.Cells | XlsSaveOptions | CachedFileFolder | temporary cache folder | large XLS export | low memory Excel save | C# Aspose.Cells example | save workbook as XLS | memory‑efficient Excel generation
// Common Searches: Aspose.Cells save large XLS without outofmemory | C# XlsSaveOptions CachedFileFolder usage | set temporary folder for Aspose.Cells export | how to reduce RAM when saving XLS in .NET | large workbook Excel 97‑2003 export Aspose
// Developer Intent: The developer needs to export a big workbook to XLS format while directing Aspose.Cells to use a temporary directory for caching, preventing excessive memory consumption.
// Use Cases: Server‑side reporting where XLS files exceed available RAM. | Web API that generates multiple large XLS documents concurrently. | Background service that batches data into XLS files without triggering OutOfMemory exceptions.
// AI Prompts: Generate C# code that saves a workbook as XLS with Aspose.Cells, specifying a custom temporary cache folder and explaining each setting. | Describe best practices for cleaning up the CachedFileFolder after saving and how to manage folder lifecycle in a web application. | Show how to tune XlsSaveOptions for optimal speed versus memory usage when exporting very large workbooks.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to configure Aspose.Cells XlsSaveOptions with the CachedFileFolder property to write a massive workbook to Excel 97‑2003 (XLS) format while keeping RAM usage low. The example creates a workbook, sets a custom temp directory for cache files, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Data");
        sheet.Cells["B1"].PutValue(12345);

        // Configure XlsSaveOptions to use a temporary folder for caching large data
        XlsSaveOptions saveOptions = new XlsSaveOptions();

        // Define a folder for temporary cache files (ensure it exists)
        string tempCacheFolder = Path.Combine(Path.GetTempPath(), "AsposeCache");
        Directory.CreateDirectory(tempCacheFolder);
        saveOptions.CachedFileFolder = tempCacheFolder;

        // Save the workbook as an Excel 97-2003 XLS file using the configured options
        workbook.Save("LargeWorkbookOutput.xls", saveOptions);
    }
}

using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Author: Example code demonstrating XLS save with temporary cache folder

        // Create a new workbook (empty or load as needed)
        Workbook workbook = new Workbook();

        // (Optional) Populate some data to illustrate usage
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample data");

        // Configure save options to use a temporary folder for caching large data
        XlsSaveOptions saveOptions = new XlsSaveOptions
        {
            // Specify a folder where Aspose.Cells can write temporary files
            CachedFileFolder = Path.Combine(Path.GetTempPath(), "AsposeCache")
        };

        // Ensure the temporary folder exists
        Directory.CreateDirectory(saveOptions.CachedFileFolder);

        // Save the workbook as an Excel 97-2003 XLS file using the configured options
        workbook.Save("output.xls", saveOptions);
    }
}
using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello OXPS");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Create XPS save options (OXPS is the same format)
        XpsSaveOptions saveOptions = new XpsSaveOptions
        {
            OnePagePerSheet = true
        };

        // Render the workbook to a memory stream using the XPS options
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Save(ms, saveOptions);
            ms.Position = 0; // Reset stream position

            // Write the OXPS bytes directly to the console output stream
            Stream consoleOut = Console.OpenStandardOutput();
            ms.CopyTo(consoleOut);
            consoleOut.Flush();
        }
    }
}
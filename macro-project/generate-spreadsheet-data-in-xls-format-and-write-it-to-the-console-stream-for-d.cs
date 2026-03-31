using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and access the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);

        // Save the workbook to a memory stream as an XLS file
        using (MemoryStream memoryStream = workbook.SaveToStream())
        {
            // Reset stream position before reading
            memoryStream.Position = 0;

            // Write the XLS bytes to the console output stream for downstream processing
            Stream consoleStream = Console.OpenStandardOutput();
            memoryStream.CopyTo(consoleStream);
            consoleStream.Flush();
        }
    }
}
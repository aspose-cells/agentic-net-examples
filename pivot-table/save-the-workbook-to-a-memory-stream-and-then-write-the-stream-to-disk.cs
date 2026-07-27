using System;
using System.IO;
using Aspose.Cells;

class SaveWorkbookToMemoryAndDisk
{
    static void Main()
    {
        // Create a new workbook (constructor rule)
        Workbook workbook = new Workbook();

        // Add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Save the workbook to a memory stream in XLSX format (Save(Stream, SaveFormat) rule)
        using (MemoryStream memoryStream = new MemoryStream())
        {
            workbook.Save(memoryStream, SaveFormat.Xlsx);
            // Reset the stream position before reading
            memoryStream.Position = 0;

            // Write the stream content to a file on disk
            using (FileStream fileStream = new FileStream("output_from_stream.xlsx", FileMode.Create, FileAccess.Write))
            {
                memoryStream.CopyTo(fileStream);
            }
        }

        // Clean up
        workbook.Dispose();
    }
}
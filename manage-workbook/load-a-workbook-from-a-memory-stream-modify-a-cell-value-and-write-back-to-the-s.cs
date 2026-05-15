using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a workbook, add initial data, and save it to a stream
        // ------------------------------------------------------------
        Workbook originalWorkbook = new Workbook();                     // new workbook
        originalWorkbook.Worksheets[0].Cells["A1"].PutValue("Original"); // initial value

        // Save the workbook to a MemoryStream (uses Workbook.SaveToStream)
        MemoryStream inputStream = originalWorkbook.SaveToStream();

        // Reset stream position so it can be read from the beginning
        inputStream.Position = 0;

        // ------------------------------------------------------------
        // 2. Load the workbook from the memory stream
        // ------------------------------------------------------------
        Workbook loadedWorkbook = new Workbook(inputStream); // loads from stream

        // ------------------------------------------------------------
        // 3. Modify a cell value
        // ------------------------------------------------------------
        loadedWorkbook.Worksheets[0].Cells["A1"].PutValue("Modified");

        // ------------------------------------------------------------
        // 4. Save the modified workbook back to a new memory stream
        // ------------------------------------------------------------
        MemoryStream outputStream = loadedWorkbook.SaveToStream(); // saves to stream

        // (Optional) Verify the modification by reading from the output stream
        outputStream.Position = 0;
        Workbook verificationWorkbook = new Workbook(outputStream);
        Console.WriteLine("Cell A1 after modification: " + verificationWorkbook.Worksheets[0].Cells["A1"].StringValue);
    }
}
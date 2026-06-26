using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a workbook and put initial data
        Workbook originalWorkbook = new Workbook();
        originalWorkbook.Worksheets[0].Cells["A1"].PutValue("Original");

        // Save the workbook to a memory stream using the provided SaveToStream method
        MemoryStream inputStream = originalWorkbook.SaveToStream();

        // Reset the stream position so it can be read from the beginning
        inputStream.Position = 0;

        // Load the workbook from the memory stream using the Stream constructor
        Workbook loadedWorkbook = new Workbook(inputStream);

        // Modify a cell value in the loaded workbook
        loadedWorkbook.Worksheets[0].Cells["A1"].PutValue("Modified");

        // Save the modified workbook back to a new memory stream
        MemoryStream outputStream = loadedWorkbook.SaveToStream();

        // Example output to verify the operation
        Console.WriteLine($"Modified stream length: {outputStream.Length} bytes");
    }
}
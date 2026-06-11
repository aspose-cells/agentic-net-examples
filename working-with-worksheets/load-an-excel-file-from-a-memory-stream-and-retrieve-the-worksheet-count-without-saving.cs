using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a workbook and add a second worksheet for demonstration
        Workbook sourceWorkbook = new Workbook();
        sourceWorkbook.Worksheets[0].Cells["A1"].PutValue("Sample");
        sourceWorkbook.Worksheets.Add();

        // Save the workbook to a memory stream using the provided SaveToStream rule
        MemoryStream memoryStream = sourceWorkbook.SaveToStream();

        // Reset the stream position before loading
        memoryStream.Position = 0;

        // Load a new workbook from the memory stream using the provided Stream constructor rule
        Workbook loadedWorkbook = new Workbook(memoryStream);

        // Retrieve the number of worksheets
        int worksheetCount = loadedWorkbook.Worksheets.Count;

        Console.WriteLine($"Worksheet count: {worksheetCount}");

        // Clean up resources
        memoryStream.Dispose();
        sourceWorkbook.Dispose();
        loadedWorkbook.Dispose();
    }
}
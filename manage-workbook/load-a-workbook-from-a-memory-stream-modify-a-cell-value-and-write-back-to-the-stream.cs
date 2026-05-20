using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // ---------- Create a new workbook ----------
        Workbook workbook = new Workbook();
        // Add initial data to cell A1
        workbook.Worksheets[0].Cells["A1"].PutValue("Original Value");

        // ---------- Save the workbook to a memory stream ----------
        // This uses the provided SaveToStream method (returns a MemoryStream)
        MemoryStream memoryStream = workbook.SaveToStream();

        // Reset the stream position to the beginning before loading
        memoryStream.Position = 0;

        // ---------- Load the workbook from the memory stream ----------
        // Uses the Workbook(Stream) constructor as per the provided rule
        Workbook loadedWorkbook = new Workbook(memoryStream);

        // ---------- Modify a cell value ----------
        loadedWorkbook.Worksheets[0].Cells["A1"].PutValue("Modified Value");

        // ---------- Save the modified workbook back to a new memory stream ----------
        MemoryStream resultStream = loadedWorkbook.SaveToStream();

        // Optional: write the result stream to a file for verification
        using (FileStream file = new FileStream("ModifiedResult.xls", FileMode.Create, FileAccess.Write))
        {
            resultStream.WriteTo(file);
        }

        // Clean up streams
        memoryStream.Dispose();
        resultStream.Dispose();
    }
}
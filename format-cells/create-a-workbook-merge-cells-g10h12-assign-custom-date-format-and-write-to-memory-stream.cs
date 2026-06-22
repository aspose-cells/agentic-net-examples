using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells G10:H12 (rows 9-11, columns 6-7)
        cells.Merge(9, 6, 3, 2);

        // Put a date value into the merged cell (top‑left corner)
        cells["G10"].PutValue(DateTime.Now);

        // Apply a custom date format to the merged cell
        Style style = cells["G10"].GetStyle();
        style.Custom = "dd-mmm-yyyy";
        cells["G10"].SetStyle(style);

        // Save the workbook to a memory stream (XLS format)
        MemoryStream stream = workbook.SaveToStream();

        // Example usage of the stream (optional)
        Console.WriteLine($"Memory stream length: {stream.Length}");
    }
}
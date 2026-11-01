// Title: Load an Excel workbook from a MemoryStream, modify cell A1, and save it back to the same MemoryStream using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells load workbook from MemoryStream and edit a cell without creating a file | How to reset a MemoryStream position before saving an updated Excel workbook with Aspose.Cells | Saving modified Excel data back to the original MemoryStream in .NET using Aspose.Cells | Update cell value in an in‑memory Excel workbook using Aspose.Cells API | Aspose.Cells example for reading and writing Excel files directly from a MemoryStream in C#
// Tags: initialize workbook from MemoryStream Aspose.Cells | change cell A1 value Aspose.Cells C# | clear and reset MemoryStream before saving Aspose.Cells | save workbook as Xlsx into MemoryStream | in‑memory Excel file manipulation Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an Excel workbook from a MemoryStream with Aspose.Cells, updating cell A1, resetting and clearing the stream, and then saving the workbook back into the same MemoryStream in XLSX format.
class Program
{
    static void Main()
    {
        // Assume inputStream already contains an Excel file.
        // In a real scenario, this stream would be provided by the caller.
        using (MemoryStream inputStream = new MemoryStream())
        {
            // Load the workbook from the memory stream.
            Workbook workbook = new Workbook(inputStream);

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Modify the value of cell A1.
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Modified Value");

            // Prepare the stream for saving: reset position and clear existing data.
            inputStream.Position = 0;
            inputStream.SetLength(0);

            // Save the updated workbook back into the same memory stream.
            workbook.Save(inputStream, SaveFormat.Xlsx);

            // The inputStream now contains the modified Excel file.
            // It can be returned, sent over a network, or written to a file as needed.
        }
    }
}

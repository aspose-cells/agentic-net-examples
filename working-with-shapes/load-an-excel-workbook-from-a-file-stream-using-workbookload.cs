// Title: Load an Excel workbook from a FileStream using Aspose.Cells in C#
// Description: Demonstrates how to open an Excel file with a FileStream, instantiate a Workbook via the Workbook(Stream) constructor, read the value of cell A1 on the first worksheet, and save the workbook to a new file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# load from stream | Workbook(Stream) constructor | read cell value Excel C# | save workbook Aspose.Cells | file stream Excel processing
// Common Searches: Aspose.Cells load workbook from FileStream | C# read Excel cell after stream load | how to save workbook after loading from stream Aspose | Workbook constructor with stream example
// Developer Intent: Open an Excel file from a stream, access cell data, and write the workbook back to disk.
// Use Cases: Process Excel files received over a network or API without creating temporary files. | Extract specific cell values for reporting or data migration after streaming the workbook. | Apply transformations (e.g., formulas, formatting) to a streamed workbook before saving the updated version.
// AI Prompts: Generate C# code that loads an Excel file from a MemoryStream with Aspose.Cells, reads several cells, and saves the result. | Show how to implement robust error handling and using‑statements when opening a workbook from a FileStream in Aspose.Cells. | Create an example that loads a workbook from a stream, inserts a shape into the first worksheet, and then saves the file.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to open an Excel file with a FileStream, instantiate a Workbook via the Workbook(Stream) constructor, read the value of cell A1 on the first worksheet, and save the workbook to a new file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Path to the Excel file that will be loaded
        string sourcePath = "input.xlsx";

        // Load the workbook from a file stream using the Workbook(Stream) constructor
        Workbook workbook;
        using (FileStream stream = new FileStream(sourcePath, FileMode.Open, FileAccess.Read))
        {
            workbook = new Workbook(stream);
        }

        // Access the first worksheet and read a sample cell value
        Worksheet sheet = workbook.Worksheets[0];
        Console.WriteLine("Value of A1: " + sheet.Cells["A1"].StringValue);

        // Save the loaded workbook to a new file (demonstrates the Save method)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}

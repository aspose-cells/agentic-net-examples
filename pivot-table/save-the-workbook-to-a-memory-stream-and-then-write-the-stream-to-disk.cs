// Title: Save an Aspose.Cells workbook to a MemoryStream and then write it as an XLSX file in C#
// Common Searches: c# aspose.cells save workbook to memory stream then to disk | how to export Aspose.Cells workbook to XLSX using MemoryStream in .NET | write Excel file from Aspose.Cells MemoryStream to file system c# | aspocells stream to file example for XLSX output
// Tags: Aspose.Cells workbook Save to MemoryStream | MemoryStream copy to FileStream Excel | C# export workbook as XLSX via stream | stream-based Excel file creation with Aspose.Cells | write XLSX from MemoryStream in .NET

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Aspose.Cells Workbook, adds sample data, saves it to a MemoryStream in XLSX format, resets the stream position, and copies the stream to a FileStream to generate an output.xlsx file on disk.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Example: add some data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Save the workbook to a memory stream in XLSX format
        using (MemoryStream memoryStream = new MemoryStream())
        {
            workbook.Save(memoryStream, SaveFormat.Xlsx);

            // Ensure the stream's position is at the beginning before writing to disk
            memoryStream.Position = 0;

            // Write the memory stream contents to a file on disk
            string outputPath = "output.xlsx";
            using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                memoryStream.CopyTo(fileStream);
            }
        }

        Console.WriteLine("Workbook saved to memory stream and written to disk successfully.");
    }
}

// Title: Save an Aspose.Cells workbook to a MemoryStream and then persist it as an XLSX file on disk using C#
// AI Prompts: Write C# code that creates an Aspose.Cells Workbook, saves it into a MemoryStream in XLSX format, resets the stream position, and copies the stream to a FileStream at a given path. | Demonstrate how to export an Aspose.Cells workbook via a MemoryStream and write the stream contents directly to a disk file without creating an intermediate workbook file.
// Common Searches: c# aspocells save workbook to memory stream then to file | how to export Aspose.Cells workbook as xlsx using MemoryStream | write Aspose.Cells workbook from MemoryStream to disk without temporary file | Aspose.Cells C# copy MemoryStream to FileStream example
// Tags: Aspose.Cells save workbook to MemoryStream | Aspose.Cells write MemoryStream to FileStream | C# export workbook as XLSX via stream | avoid temporary file Aspose.Cells save | MemoryStream to disk Aspose.Cells

using Aspose.Cells;
using System.IO;

// The program creates a new Aspose.Cells workbook, adds sample data, saves it to a MemoryStream in XLSX format, resets the stream position, and then copies the stream contents to a file named output.xlsx on the local disk.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Add sample data to the first worksheet
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Save the workbook to a memory stream in XLSX format
        using (var memoryStream = new MemoryStream())
        {
            workbook.Save(memoryStream, SaveFormat.Xlsx);

            // Reset the stream position to the beginning before reading
            memoryStream.Position = 0;

            // Write the memory stream contents to a file on disk
            using (var fileStream = new FileStream("output.xlsx", FileMode.Create, FileAccess.Write))
            {
                memoryStream.CopyTo(fileStream);
            }
        }
    }
}

// Title: Convert a large HTML file to XLSX using Aspose.Cells with MemoryStream in C#
// Common Searches: how to load HTML into an Aspose.Cells workbook from a MemoryStream in C# | convert HTML to XLSX without creating temporary files using Aspose.Cells | process large HTML content in memory and export to Excel with Aspose.Cells C# | save an Aspose.Cells workbook to a MemoryStream as XLSX in C# | use LoadOptions to import HTML format into an Aspose.Cells workbook
// Tags: Aspose.Cells load HTML from MemoryStream | convert HTML to XLSX C# | save workbook to MemoryStream Aspose.Cells | process large HTML without temp files Aspose.Cells | LoadOptions HTML format Aspose.Cells | in‑memory HTML to Excel conversion C#

using System;
using System.IO;
using Aspose.Cells;

// The example reads a large HTML file into a MemoryStream, loads it into an Aspose.Cells Workbook using LoadOptions for HTML format, then saves the workbook to another MemoryStream as an XLSX file, allowing the Excel data to be used without creating intermediate temporary files.
class Program
{
    static void Main()
    {
        // Example: read a large HTML file into a memory stream.
        // In real scenarios the HTML could come from a network stream, database, etc.
        using (FileStream htmlFileStream = File.OpenRead("input.html"))
        using (MemoryStream htmlMemoryStream = new MemoryStream())
        {
            // Copy the source HTML into the memory stream.
            // This allows Aspose.Cells to work with the data without creating temporary files.
            htmlFileStream.CopyTo(htmlMemoryStream);
            htmlMemoryStream.Position = 0; // Reset position for reading.

            // Load the HTML content into an Aspose.Cells Workbook.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(htmlMemoryStream, loadOptions);

            // Save the workbook to another memory stream in XLSX format.
            using (MemoryStream excelMemoryStream = new MemoryStream())
            {
                workbook.Save(excelMemoryStream, SaveFormat.Xlsx);
                // At this point excelMemoryStream contains the Excel file.
                // It can be sent over a network, stored in a database, or written to disk.
                File.WriteAllBytes("output.xlsx", excelMemoryStream.ToArray());
            }
        }
    }
}

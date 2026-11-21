// Title: How to add a hyperlink to an embedded OLE object that opens the original source file using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add hyperlink to an embedded OLE object | How to open the original DOCX from an OLE object in Excel using Aspose.Cells | Create an OLE object with a link to the source file in Aspose.Cells for .NET | C# example of embedding a DOCX as OLE and adding a hyperlink in Excel | Add a hyperlink to the cell that contains an OLE object with Aspose.Cells
// Tags: embed docx as OLE object Aspose.Cells | OleObjects.Add embed file Aspose.Cells | hyperlink to source file Aspose.Cells | cell hyperlink for OLE object Aspose.Cells | C# Excel OLE embedding Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, embeds a DOCX file as an OLE object at cell B2, adds a hyperlink on the same cell that points to the original source file path, and saves the workbook as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the original source file that will be embedded as OLE object
            string sourceFilePath = @"C:\Temp\SampleDocument.docx";

            // Verify that the source file exists
            if (!File.Exists(sourceFilePath))
                throw new FileNotFoundException("Source file not found.", sourceFilePath);

            // Read the source file into a byte array (required by OleObjects.Add)
            byte[] sourceFileBytes = File.ReadAllBytes(sourceFilePath);

            // Insert the OLE object at cell B2 (row 1, column 1) with specified size
            int oleIndex = sheet.OleObjects.Add(1, 1, 100, 200, sourceFileBytes);
            OleObject ole = sheet.OleObjects[oleIndex];

            // Add a hyperlink to the same cell that points to the original source file
            // Note: Hyperlinks.Add(string hyperlink, int firstRow, int firstColumn, string screenTip)
            sheet.Hyperlinks.Add(sourceFilePath, 1, 1, "Open original source file");

            // Ensure the output directory exists
            string outputPath = @"C:\Temp\HyperlinkedOleObject.xlsx";
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook to a file
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

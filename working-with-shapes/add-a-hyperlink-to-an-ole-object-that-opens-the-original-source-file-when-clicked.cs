// Title: Generate an Excel file with an embedded PDF OLE object that opens the original PDF via a file hyperlink using Aspose.Cells for .NET
// AI Prompts: Write C# code that reads a PDF file, inserts it as an OLE object into cell B2 of a new workbook, sets the OleObject.Hyperlink.Address to a file:// URL of the PDF, and saves the workbook. | Show how to adjust the width and height of an OLE object after embedding a PDF and assign a hyperlink that points to the source document using Aspose.Cells. | Demonstrate error handling for missing source files when adding a PDF OLE object with a hyperlink in an Aspose.Cells worksheet.
// Common Searches: aspnet embed pdf as ole object in excel and link to original file | c# aspose.cells add hyperlink to oleobject pointing to local file | how to set file:// hyperlink for embedded pdf in excel using aspose cells | save workbook with pdf ole object and clickable link c# | aspose.cells oleobject hyperlink address property example
// Tags: Aspose.Cells embed PDF OLE object | OleObject hyperlink file URL | C# set OleObject.Hyperlink.Address | Excel OLE object size adjustment Aspose | handle missing source file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, embeds a PDF file as an OLE object in cell B2, assigns a file:// hyperlink that opens the original PDF, optionally resizes the object, handles missing source files, and saves the workbook to the specified location.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source PDF file
            string sourceFilePath = @"C:\Files\Document.pdf";
            if (!File.Exists(sourceFilePath))
                throw new FileNotFoundException("Source file not found.", sourceFilePath);

            // Path for the output workbook
            string outputPath = @"C:\Output\WorkbookWithOleHyperlink.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Read the source file into a byte array (required by OleObjects.Add)
            byte[] oleData = File.ReadAllBytes(sourceFilePath);

            // Add the OLE object (PDF) to cell B2 (row 1, column 1) with initial size
            // Add returns the index of the newly added object
            int oleIndex = sheet.OleObjects.Add(1, 1, 150, 200, oleData);
            OleObject ole = sheet.OleObjects[oleIndex];

            // Assign a hyperlink that points to the original file
            ole.Hyperlink.Address = $"file:///{sourceFilePath.Replace("\\", "/")}";

            // Optionally adjust the size of the OLE object (in points)
            ole.Width = 200;
            ole.Height = 150;

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

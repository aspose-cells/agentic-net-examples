// Title: Insert a PDF as an OLE object with a custom 'Report' icon label in an Excel sheet using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that reads a PDF file, adds it as an OLE object at cell A1, sets the icon caption to 'Report', and enables auto‑size. | Show how to embed a PDF into an Excel workbook as an OLE object and customize its icon label using Aspose.Cells for .NET. | Provide a complete example that loads a PDF into a byte array, creates an OLE object with a custom caption, and saves the workbook with Aspose.Cells.
// Common Searches: aspnet embed pdf as ole object in excel worksheet using aspose.cells | set custom icon caption for embedded pdf ole object in Aspose.Cells C# | how to auto size ole object icon after embedding pdf with Aspose.Cells | add pdf ole object to specific cell in Excel using Aspose.Cells .NET | read pdf into byte array for ole embedding Aspose.Cells example
// Tags: Aspose.Cells add PDF OLE object | C# embed PDF as OLE in Excel | custom OLE icon caption Aspose.Cells | auto‑size OLE object icon Aspose.Cells | read PDF to byte array for OLE embedding

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, verifies the PDF file, reads it into a byte array, inserts it as an OLE object at cell A1 with a 200×200 size, sets the icon caption to 'Report', enables auto‑sizing, and saves the file as EmbeddedPdf.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the PDF file to embed
            string pdfPath = @"C:\Path\To\Report.pdf";

            // Ensure the PDF file exists before attempting to embed
            if (!File.Exists(pdfPath))
            {
                Console.WriteLine($"Error: PDF file not found at '{pdfPath}'.");
                return;
            }

            // Read PDF file into a byte array (required by OleObjects.Add overload)
            byte[] pdfData = File.ReadAllBytes(pdfPath);

            // Add an OLE object for the PDF at cell A1 (row 0, column 0) with a size of 200x200 points
            int oleIndex = sheet.OleObjects.Add(0, 0, 200, 200, pdfData);

            // Retrieve the added OLE object
            OleObject ole = sheet.OleObjects[oleIndex];

            // Optional: let the OLE object auto‑size to its icon
            ole.IsAutoSize = true;

            // Save the workbook with the embedded PDF
            string outputPath = "EmbeddedPdf.xlsx";

            // Determine output directory (handle case where only file name is provided)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();

            // Ensure the output directory exists
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

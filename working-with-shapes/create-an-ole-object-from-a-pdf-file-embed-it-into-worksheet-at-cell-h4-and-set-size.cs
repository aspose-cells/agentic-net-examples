// Title: C# – Embed a PDF as an OLE object in cell H4 using Aspose.Cells
// Description: Load a PDF file, create a tiny PNG icon, add an OLE object at cell H4 (200 × 200), set ProgID to AcroExch.Document.DC, display it as an icon with a label, and save the workbook with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | embed PDF OLE object | Excel OLE object cell H4 | set OLE size Aspose.Cells | display as icon | ProgID AcroExch.Document.DC | binary data embedding | sample code | GitHub example
// Common Searches: Aspose.Cells embed PDF as OLE object | Add OLE object to specific Excel cell C# | Set OLE object dimensions in Aspose.Cells | Display PDF OLE as icon with label | ProgID for PDF OLE in Aspose.Cells
// Developer Intent: Insert a PDF file into an Excel worksheet as an OLE object positioned at cell H4, define its size, and show it as a clickable icon.
// Use Cases: Generate a report that bundles supporting PDFs directly inside the spreadsheet. | Create a template where users can open detailed specifications by double‑clicking an icon in a designated cell. | Automate packaging of multiple PDFs into a single Excel file for easy distribution to clients.
// AI Prompts: Write C# code with Aspose.Cells to embed a PDF as an OLE object at cell H4, size 200x200, displayed as an icon labeled "Sample PDF". | Explain why the ProgID "AcroExch.Document.DC" is required for PDF OLE objects in Aspose.Cells and how it affects opening the file. | Provide best‑practice error handling for reading binary files and adding OLE objects with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Load a PDF file, create a tiny PNG icon, add an OLE object at cell H4 (200 × 200), set ProgID to AcroExch.Document.DC, display it as an icon with a label, and save the workbook with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Verify that the PDF file exists before attempting to read it
            const string pdfPath = "sample.pdf";
            if (!File.Exists(pdfPath))
            {
                Console.WriteLine($"Error: PDF file \"{pdfPath}\" not found.");
                return;
            }

            // Load PDF binary data
            byte[] pdfData = File.ReadAllBytes(pdfPath);

            // Minimal 1x1 PNG icon (transparent) for the OLE object
            byte[] iconData = new byte[]
            {
                0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
                0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
                0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
                0x54,0x78,0x9C,0x63,0x00,0x01,0x00,0x00,
                0x05,0x00,0x01,0x0D,0x0A,0x2D,0xB4,0x00,
                0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
                0x42,0x60,0x82
            };

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add an OLE object at cell H4 (row index 3, column index 7) with the generated icon
            int oleIndex = worksheet.OleObjects.Add(3, 7, 200, 200, iconData);
            OleObject oleObject = worksheet.OleObjects[oleIndex];

            // Assign the PDF binary data to the OLE object
            oleObject.ObjectData = pdfData;

            // Set the ProgID for PDF files (AcroExch.Document.DC)
            oleObject.ProgID = "AcroExch.Document.DC";

            // Display the OLE object as an icon and set a label
            oleObject.DisplayAsIcon = true;
            oleObject.Label = "Sample PDF";

            // Ensure output directory exists
            const string outputPath = "OutputWithPdfOle.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the embedded PDF OLE object
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}

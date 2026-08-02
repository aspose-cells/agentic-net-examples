// Title: C# – Embed a PDF as an OLE object with a custom “Report” icon caption using Aspose.Cells
// Description: Shows how to check for a PDF file, read its bytes, create a workbook, add an OLE placeholder with a transparent PNG, embed the PDF as an embedded object, display it as an icon labeled “Report”, and save the sheet as an XLSX file.
// Keywords: Aspose.Cells embed PDF | C# OLE object Excel | SetEmbeddedObject Aspose | custom icon caption | display PDF as icon | Excel workbook OLE | embedded PDF worksheet
// Common Searches: Aspose.Cells embed PDF as OLE object C# | set custom icon caption for OLE object in Excel using Aspose | display embedded PDF as clickable icon in .NET workbook | how to add PDF OLE object without linking file Aspose.Cells | change OLE object label in Aspose.Cells
// Developer Intent: Insert a PDF into an Excel worksheet as an embedded OLE object, show it as an icon, and assign the caption "Report".
// Use Cases: Distribute a spreadsheet that contains attached PDF documentation accessible via an icon. | Create a financial report workbook where supporting PDFs are embedded for one‑click access. | Generate a printable Excel file that includes embedded policy PDFs without external dependencies.
// AI Prompts: Write C# code with Aspose.Cells to embed a PDF as an OLE object, use a custom PNG icon, and set the label to "Report" at cell F6. | Explain how to modify the size and position of the PDF OLE icon after it is added to a worksheet. | Provide a method to extract or replace the embedded PDF in an existing Aspose.Cells workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to check for a PDF file, read its bytes, create a workbook, add an OLE placeholder with a transparent PNG, embed the PDF as an embedded object, display it as an icon labeled “Report”, and save the sheet as an XLSX file.
class EmbedPdfOleObject
{
    static void Main()
    {
        try
        {
            // Verify PDF file exists
            string pdfPath = "sample.pdf";
            if (!File.Exists(pdfPath))
                throw new FileNotFoundException("PDF file not found.", pdfPath);

            // Read PDF bytes
            byte[] pdfBytes = File.ReadAllBytes(pdfPath);

            // 1x1 transparent PNG (used as the OLE icon)
            byte[] iconBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");

            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            // Add an OLE object placeholder with the icon image
            int oleIndex = worksheet.OleObjects.Add(5, 2, 150, 150, iconBytes);
            OleObject oleObject = worksheet.OleObjects[oleIndex];

            // Embed the PDF, display as an icon with caption "Report"
            bool linkToFile = false;   // embed the file, not link
            bool displayAsIcon = true; // show as icon
            string label = "Report";   // icon caption
            oleObject.SetEmbeddedObject(linkToFile, pdfBytes, Path.GetFileName(pdfPath), displayAsIcon, label);

            // Ensure the object is displayed as an icon with the correct label
            oleObject.DisplayAsIcon = true;
            oleObject.Label = label;

            // Save the workbook
            string outputPath = "WorkbookWithPdfOleObject.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"File error: {ex.Message}");
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

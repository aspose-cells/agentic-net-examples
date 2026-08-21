// Title: Embed a PDF as an OLE icon in an Excel workbook with Aspose.Cells for .NET
// Description: C# example that loads or creates a workbook, reads a PDF file, inserts it as an OLE object at a specified cell range, sets the ProgID to Acrobat, configures the object to display as an icon with a filename label, and saves the workbook.
// Keywords: Aspose.Cells | .NET | C# | embed PDF | OLE object | display as icon | AcroExch.Document.DC | Excel workbook | add OLE object programmatically | PDF attachment in Excel
// Common Searches: Aspose.Cells embed PDF as OLE icon | C# add PDF OLE object to Excel | display PDF as icon in Excel using Aspose | set ProgID for PDF OLE object Aspose.Cells | how to insert PDF into worksheet as icon
// Developer Intent: Insert a PDF file into a worksheet as an OLE object shown as an icon and save the file.
// Use Cases: Create a report that bundles supporting PDFs as clickable icons. | Distribute a single Excel package containing multiple PDF manuals. | Design a template where users can view PDF references without opening separate files.
// AI Prompts: Generate C# code with Aspose.Cells to embed a PDF as an OLE object, show it as an icon, and label it with the file name. | Adapt the example to place the OLE object using a cell address like "E5" and scale its size automatically. | Add robust error handling that logs missing PDF files and continues processing remaining worksheets.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that loads or creates a workbook, reads a PDF file, inserts it as an OLE object at a specified cell range, sets the ProgID to Acrobat, configures the object to display as an icon with a filename label, and saves the workbook.
class OleObjectPdfDemo
{
    static void Main()
    {
        try
        {
            // Load an existing workbook if it exists; otherwise create a new one.
            string workbookPath = "input.xlsx";
            Workbook workbook;
            if (File.Exists(workbookPath))
            {
                workbook = new Workbook(workbookPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Read PDF file bytes if the file exists
            string pdfPath = "sample.pdf";
            byte[] pdfBytes = File.Exists(pdfPath) ? File.ReadAllBytes(pdfPath) : Array.Empty<byte>();

            // Add an OLE object placeholder (image data can be empty)
            int topRow = 5;          // upper left row index
            int leftColumn = 5;      // upper left column index
            int height = 200;        // height in pixels
            int width = 200;         // width in pixels
            int oleIndex = worksheet.OleObjects.Add(topRow, leftColumn, height, width, new byte[0]);

            // Configure the OLE object
            OleObject oleObject = worksheet.OleObjects[oleIndex];
            oleObject.ObjectData = pdfBytes;                     // embed PDF data
            oleObject.ProgID = "AcroExch.Document.DC";           // PDF ProgID
            oleObject.DisplayAsIcon = true;                     // show as icon
            oleObject.Label = Path.GetFileName(pdfPath);        // icon label

            // Save the workbook with the embedded PDF OLE object
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Title: C# – Embed a PDF as an OLE icon in an Excel workbook using Aspose.Cells for .NET
// Description: Loads an existing .xlsx file, reads a PDF into a byte array, inserts the PDF as an OLE object at cell F6 (row 6, column 3) with a 200 × 200 size, sets the ProgID to AcroExch.Document.DC, enables DisplayAsIcon, and saves the updated workbook.
// Keywords: Aspose.Cells | C# | .NET | embed PDF | OLE object | DisplayAsIcon | ProgID | Excel worksheet | add OLE to cell | sample code
// Common Searches: Aspose.Cells add PDF as OLE object | C# embed PDF in Excel with icon | Display OLE object as icon Aspose.Cells | Set ProgID for PDF OLE in .NET | Insert OLE object into specific Excel cell
// Developer Intent: Insert a PDF into an Excel sheet as an OLE object displayed as an icon and save the workbook.
// Use Cases: Create a report that bundles supporting PDFs accessible via clickable icons. | Distribute a single Excel file containing multiple PDF attachments for stakeholder review. | Build a template where users can open embedded documentation by double‑clicking an icon.
// AI Prompts: Generate C# code with Aspose.Cells to embed a PDF at row 6 column 3 and show it as an icon. | Explain how to choose the correct ProgID for different file types when adding OLE objects in Aspose.Cells. | Suggest best practices for handling large PDF files as OLE objects in an Excel workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an existing .xlsx file, reads a PDF into a byte array, inserts the PDF as an OLE object at cell F6 (row 6, column 3) with a 200 × 200 size, sets the ProgID to AcroExch.Document.DC, enables DisplayAsIcon, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Read the PDF file to embed
        byte[] pdfData = File.ReadAllBytes("sample.pdf");

        // Add the PDF as an OLE object (uses OleObjectCollection.Add rule)
        int oleIndex = sheet.OleObjects.Add(5, 2, 200, 200, pdfData);

        // Retrieve the added OLE object
        OleObject oleObject = sheet.OleObjects[oleIndex];

        // Set the ProgID for PDF files
        oleObject.ProgID = "AcroExch.Document.DC";

        // Display the OLE object as an icon (uses DisplayAsIcon property rule)
        oleObject.DisplayAsIcon = true;

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}

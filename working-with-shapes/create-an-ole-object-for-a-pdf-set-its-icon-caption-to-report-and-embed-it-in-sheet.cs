// Title: Add a PDF OLE object with a custom icon and caption to an Excel worksheet using Aspose.Cells for .NET
// Common Searches: how to embed a PDF as an OLE object with a custom icon in an Excel file using Aspose.Cells C# | Aspose.Cells C# set OLE object label and display as icon for PDF | create OLE object from PDF bytes and custom PNG icon in Aspose.Cells workbook | set ProgID for PDF OLE object when adding to worksheet with Aspose.Cells | save workbook with embedded PDF OLE object using Aspose.Cells .NET
// Tags: PDF OLE embedding with custom icon Aspose.Cells | set OLE object caption in Excel using Aspose.Cells | display OLE as icon in .NET workbook | embed PDF byte array into worksheet OLE object | assign ProgID for PDF OLE in Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a PDF file and a PNG icon into byte arrays, creates a new Workbook, adds an OLE object at a specific cell using the icon bytes, assigns the PDF bytes as the object's data, sets the file format to PDF, enables display as an icon, defines the caption "Report", optionally sets the ProgID and image source path, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        // Paths to the source PDF, the icon image and the output workbook.
        // Adjust these paths to point to existing files on your machine.
        string pdfPath = "sample.pdf";
        string iconPath = "icon.png";
        string outputPath = "WorkbookWithOle.xlsx";

        // Verify that the required files exist.
        if (!File.Exists(pdfPath))
        {
            Console.WriteLine($"Error: PDF file not found at '{pdfPath}'.");
            return;
        }

        if (!File.Exists(iconPath))
        {
            Console.WriteLine($"Error: Icon image file not found at '{iconPath}'.");
            return;
        }

        try
        {
            // Load the PDF and icon image into byte arrays.
            byte[] pdfBytes = File.ReadAllBytes(pdfPath);
            byte[] iconBytes = File.ReadAllBytes(iconPath);

            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add an OLE object to the worksheet using the icon image.
            // Parameters: upperLeftRow, upperLeftColumn, width (pixels), height (pixels), image data.
            int oleIndex = sheet.OleObjects.Add(2, 2, 100, 100, iconBytes);
            OleObject oleObject = sheet.OleObjects[oleIndex];

            // Embed the PDF data.
            oleObject.ObjectData = pdfBytes;

            // Set the file format type to PDF.
            oleObject.FileFormatType = FileFormatType.Pdf;

            // Display the OLE object as an icon.
            oleObject.DisplayAsIcon = true;

            // Set the caption (label) that appears under the icon.
            oleObject.Label = "Report";

            // Optional: set the ProgID for PDF files (helps Excel choose the correct application).
            oleObject.ProgID = "AcroExch.Document.DC";

            // Optional: set the full path of the icon image (used by Excel when the file is opened).
            oleObject.ImageSourceFullName = iconPath;

            // Save the workbook.
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}

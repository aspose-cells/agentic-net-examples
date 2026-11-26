// Title: How to embed a PDF file as an OLE object with a custom icon in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# embed PDF as OLE object and show as icon in Excel | Add OLE object with custom icon to worksheet using Aspose.Cells .NET | Set DisplayAsIcon property for PDF OLE object in Aspose.Cells workbook | Save workbook with embedded PDF OLE object using Aspose.Cells C# | Load PDF bytes and icon image into OLE object in Aspose.Cells
// Tags: add OLE object with custom icon Aspose.Cells | embed PDF as OLE object Excel .NET | set OLE object DisplayAsIcon property C# | assign ObjectData bytes to OLE object Aspose.Cells | specify FileFormatType Pdf for OLE object

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectPdfDemo
{
    // The example reads a PDF file and an icon image, creates a new workbook, adds an OLE object at a specific cell using the icon bytes, assigns the PDF bytes to the object's data, sets the file format to PDF, enables display as an icon, labels the object with the PDF filename, and saves the workbook as an .xlsx file.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths – adjust if necessary. Files must exist in the same folder as the executable
            const string pdfPath = "sample.pdf";
            const string iconPath = "icon.png";
            const string outputPath = "WorkbookWithPdfOle.xlsx";

            try
            {
                // Verify input files
                if (!File.Exists(pdfPath))
                    throw new FileNotFoundException($"PDF file not found: {pdfPath}");
                if (!File.Exists(iconPath))
                    throw new FileNotFoundException($"Icon image file not found: {iconPath}");

                // Load PDF bytes (the data that will be embedded)
                byte[] pdfBytes = File.ReadAllBytes(pdfPath);

                // Load icon image bytes (the picture shown in the worksheet)
                byte[] iconBytes = File.ReadAllBytes(iconPath);

                // Create a new workbook (empty)
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add an OLE object to the sheet.
                // Parameters: upper left row, upper left column, height (pixels), width (pixels), image data (icon)
                int oleIndex = sheet.OleObjects.Add(2, 2, 100, 100, iconBytes);

                // Retrieve the newly added OleObject
                OleObject oleObject = sheet.OleObjects[oleIndex];

                // Set the embedded PDF data
                oleObject.ObjectData = pdfBytes;

                // Specify that the embedded object is a PDF file
                oleObject.FileFormatType = FileFormatType.Pdf;

                // Display the OLE object as an icon
                oleObject.DisplayAsIcon = true;

                // Optional: set a friendly label that appears when the icon is hovered
                oleObject.Label = Path.GetFileName(pdfPath);

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("The PDF is embedded as an OLE object displayed as an icon.");
            }
            catch (Exception ex)
            {
                // Do not hide failures – report the full exception
                Console.Error.WriteLine("An error occurred:");
                Console.Error.WriteLine(ex.ToString());
            }
        }
    }
}

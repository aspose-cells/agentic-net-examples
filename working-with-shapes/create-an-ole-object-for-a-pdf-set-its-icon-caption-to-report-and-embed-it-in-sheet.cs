// Title: How to embed a PDF as an OLE object with a custom icon caption in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# embed PDF as OLE object with custom icon caption | Add PDF OLE object to Excel sheet programmatically using Aspose.Cells | Set icon caption for OLE object in Aspose.Cells workbook | Create OLE object from PDF file in Aspose.Cells .NET example | Display PDF OLE object as icon in Excel using Aspose.Cells
// Tags: embed pdf ole object Aspose.Cells | custom ole icon caption Aspose.Cells | add ole object to worksheet C# | oleobjects.add pdf parameters | display ole object as icon Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, reads a PDF file into a byte array, adds it as an OLE object at cell B2 with a 100 × 100 px size and the icon caption “Report”, optionally sets it to display as an icon, and saves the workbook as WorkbookWithPdfOle.xlsx.
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

            // Path to the PDF file that will be embedded as an OLE object
            string pdfFilePath = "Report.pdf";

            // Ensure the PDF file exists before proceeding
            if (!File.Exists(pdfFilePath))
            {
                Console.WriteLine($"File not found: {pdfFilePath}");
                return;
            }

            // Read the PDF file into a byte array (required by the Add method)
            byte[] pdfData = File.ReadAllBytes(pdfFilePath);

            // Add the OLE object to the worksheet.
            // Parameters: upper left row, upper left column, height (pixels), width (pixels), file data, icon caption
            int oleIndex = sheet.OleObjects.Add(1, 1, 100, 100, pdfData, "Report");

            // Retrieve the added OleObject
            OleObject oleObject = sheet.OleObjects[oleIndex];

            // If the API version supports it, display the OLE object as an icon.
            // Uncomment the following line if IsObjectIcon is available in your Aspose.Cells version.
            // oleObject.IsObjectIcon = true;

            // Prepare output path
            string outputPath = "WorkbookWithPdfOle.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

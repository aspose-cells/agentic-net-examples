// Title: Create an Excel hyperlink to a PDF on a UNC network share using Aspose.Cells for .NET (C#)
// Common Searches: how to insert a UNC path hyperlink to a PDF in an Excel file using Aspose.Cells C# | Aspose.Cells add hyperlink to network share PDF example | C# code to link a worksheet cell to a PDF located on a file server with Aspose.Cells | using Aspose.Cells Hyperlinks.Add to reference a PDF on a Windows share | save workbook after adding external PDF link with Aspose.Cells .NET
// Tags: Aspose.Cells Hyperlinks.Add UNC path | add external PDF hyperlink Excel C# | link to network share PDF Aspose.Cells | set TextToDisplay for worksheet hyperlink | save workbook after adding hyperlink Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, writes "Open PDF" to cell B2, defines a UNC path to a PDF on a network share, adds a hyperlink to that cell using Hyperlinks.Add, sets the display text, ensures the output directory exists, and saves the workbook as HyperlinkExample.xlsx while handling potential errors.
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

            // Define the cell where the hyperlink will be placed
            Cell cell = sheet.Cells["B2"];
            cell.PutValue("Open PDF");

            // Path to the external PDF on a network share
            // Example: \\ServerName\ShareFolder\Document.pdf
            string pdfPath = @"\\ServerName\ShareFolder\Document.pdf";

            // Ensure the PDF path is not empty (actual existence check is optional for network shares)
            if (string.IsNullOrWhiteSpace(pdfPath))
                throw new FileNotFoundException("PDF path is not specified.");

            // Add a hyperlink to the cell that points to the PDF.
            // Hyperlinks.Add(row, column, totalRows, totalColumns, hyperlinkAddress)
            // Row and column indices are zero‑based; B2 => row 1, column 1.
            int hyperlinkIndex = sheet.Hyperlinks.Add(1, 1, 1, 1, pdfPath);

            // Set the display text for the hyperlink
            sheet.Hyperlinks[hyperlinkIndex].TextToDisplay = "Open PDF";

            // Save the workbook to a file
            string outputPath = "HyperlinkExample.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

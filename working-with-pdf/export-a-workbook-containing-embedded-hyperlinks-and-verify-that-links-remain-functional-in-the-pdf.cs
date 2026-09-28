// Title: Save an Excel workbook with a cell hyperlink as a PDF while keeping the link clickable using Aspose.Cells for .NET
// AI Prompts: Generate C# code that inserts a hyperlink into cell A1 of a new workbook and exports it to PDF, preserving the clickable link with Aspose.Cells. | Write a C# snippet that checks for a specific URL in a worksheet's Hyperlinks collection before converting the workbook to PDF using Aspose.Cells. | Create C# logic that ensures the PDF output folder exists, saves the workbook as a PDF, and gracefully handles any exceptions while maintaining Excel hyperlinks.
// Common Searches: Aspose.Cells keep hyperlinks clickable when converting Excel to PDF in C# | C# export workbook to PDF with active hyperlinks using Aspose.Cells | how to verify an Excel hyperlink exists before saving as PDF with Aspose.Cells | create PDF from Excel with embedded links and ensure output directory exists .NET
// Tags: Aspose.Cells add hyperlink to cell | export workbook to PDF with Aspose.Cells | preserve Excel hyperlinks in PDF | C# verify worksheet hyperlinks | handle output directory Aspose.Cells PDF

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, names the first worksheet, adds display text and a hyperlink to cell A1, ensures the target PDF directory exists, saves the workbook as a PDF using Aspose.Cells, iterates through the worksheet's Hyperlinks collection to confirm the URL is present, and outputs a verification message while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Links";

            // Add display text to cell A1
            sheet.Cells["A1"].PutValue("Visit OpenAI");

            // Add a hyperlink to cell A1 (row 0, column 0)
            // totalRows = 1, totalColumns = 1 for a single cell hyperlink
            sheet.Hyperlinks.Add(0, 0, 1, 1, "https://www.openai.com");

            // Define PDF output path
            string pdfPath = "Hyperlinks.pdf";

            // Ensure the output directory exists
            string pdfDir = Path.GetDirectoryName(Path.GetFullPath(pdfPath));
            if (!string.IsNullOrEmpty(pdfDir) && !Directory.Exists(pdfDir))
            {
                Directory.CreateDirectory(pdfDir);
            }

            // Save the workbook as PDF (Aspose.Cells creates the file)
            workbook.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the hyperlink exists in the worksheet
            bool hyperlinkFound = false;
            foreach (Hyperlink hl in sheet.Hyperlinks)
            {
                if (hl.Address == "https://www.openai.com")
                {
                    hyperlinkFound = true;
                    break;
                }
            }

            Console.WriteLine(hyperlinkFound
                ? "Hyperlink verified in workbook and PDF saved."
                : "Hyperlink not found in workbook.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

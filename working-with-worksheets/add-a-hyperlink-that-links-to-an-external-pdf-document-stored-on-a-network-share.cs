// Title: Insert a hyperlink to a PDF on a UNC network share into an Excel cell with Aspose.Cells for .NET (C#)
// AI Prompts: Create a new workbook, write display text to a cell, and add a file:// hyperlink that points to a PDF located on a UNC path using Aspose.Cells. | Check that a network‑share PDF file exists, convert the UNC path to a URI, and attach the hyperlink to cell A1 with Aspose.Cells. | Save the workbook after embedding the external PDF link and output the full path of the generated .xlsx file.
// Common Searches: Aspose.Cells C# add hyperlink to PDF on network share UNC path | How to create file:// link to a shared PDF in Excel using Aspose.Cells | Validate existence of a PDF on a network share before inserting hyperlink with Aspose.Cells | Convert UNC path to URI for Excel hyperlink in C# Aspose.Cells example | Add external document hyperlink to Excel cell programmatically with Aspose.Cells .NET
// Tags: Aspose.Cells add UNC hyperlink | Excel file URI hyperlink C# | network share PDF hyperlink Aspose | validate file existence before hyperlink Aspose.Cells | hyperlink cell A1 Aspose.Cells example

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a new workbook, writes "Open PDF Document" to cell A1, verifies that a PDF file exists on a UNC network share, converts the UNC path to a file:// URI, adds the hyperlink to the cell, and saves the workbook as HyperlinkExample.xlsx.
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

            // Set the display text for the hyperlink in cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Open PDF Document");

            // UNC path to the PDF on the network share
            string pdfPath = @"\\ServerName\SharedFolder\Document.pdf";

            // Validate the UNC path
            if (string.IsNullOrWhiteSpace(pdfPath))
                throw new FileNotFoundException("The PDF path is not specified.");

            // Optional: verify the file exists to avoid runtime errors
            if (!File.Exists(pdfPath))
                throw new FileNotFoundException($"The PDF file was not found at '{pdfPath}'.");

            // Convert the UNC path to a file URI (e.g., file://ServerName/SharedFolder/Document.pdf)
            string fileUri = new Uri(pdfPath).AbsoluteUri;

            // Add the hyperlink to cell A1 (row 0, column 0, 1 row, 1 column)
            sheet.Hyperlinks.Add(0, 0, 1, 1, fileUri);

            // Define output file name
            string outputPath = "HyperlinkExample.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

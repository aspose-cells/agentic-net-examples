// Title: C# – Insert Column at Index 3 and Export Worksheet to PDF with Aspose.Cells
// Description: Load an existing Excel file using Aspose.Cells for .NET, insert a new column before column D (zero‑based index 3) in the first worksheet, and save the updated workbook as a PDF. The snippet includes file‑existence validation and basic exception handling.
// Keywords: Aspose.Cells insert column C# | add column index 3 Aspose.Cells | export worksheet to PDF C# | Aspose.Cells PDF conversion example | C# Excel to PDF Aspose.Cells | insert column before D Excel .NET | Aspose.Cells try‑catch sample | GitHub Aspose.Cells column insert
// Common Searches: How to insert a column at a specific index with Aspose.Cells and save as PDF in C# | C# code to add a blank column before column D and export workbook to PDF | Aspose.Cells example: insert column and convert Excel to PDF | Insert column at index 3 Aspose.Cells .NET tutorial | Export modified worksheet to PDF using Aspose.Cells
// Developer Intent: Add a column at position 3 in an Excel workbook and generate a PDF of the modified sheet using Aspose.Cells for .NET.
// Use Cases: Prepare a financial statement by inserting a placeholder column before generating a client‑ready PDF. | Update a reporting template with an extra metric column, then export the sheet for printing or distribution. | Automate batch processing where each workbook needs a new column added prior to PDF conversion.
// AI Prompts: Write C# code that uses Aspose.Cells to insert a column at index 3 in the first worksheet and save the workbook as a PDF with custom page margins. | Provide an Aspose.Cells example with try‑catch blocks, file‑existence checks, column insertion, and PDF export. | Explain how to configure PdfSaveOptions (orientation, compression, header/footer) when exporting a workbook after inserting a column with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Load an existing Excel file using Aspose.Cells for .NET, insert a new column before column D (zero‑based index 3) in the first worksheet, and save the updated workbook as a PDF. The snippet includes file‑existence validation and basic exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the existing Excel file
            string inputFile = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file \"{inputFile}\" was not found.");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputFile);

            // Get the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Insert a new column at index 3 (zero‑based, before column D)
            worksheet.Cells.InsertColumn(3);

            // Prepare PDF save options (optional customizations can be set here)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Export the modified worksheet to a PDF file
            workbook.Save("output.pdf", pdfOptions);

            Console.WriteLine("Workbook successfully saved as PDF.");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

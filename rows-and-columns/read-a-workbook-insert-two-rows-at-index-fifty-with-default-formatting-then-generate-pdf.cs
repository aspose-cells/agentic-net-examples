// Title: Add two blank rows at row index 50 in an Excel worksheet using Aspose.Cells for .NET and convert the workbook to PDF
// AI Prompts: Add two blank rows at row index 50 in an existing .xlsx file with Aspose.Cells for .NET, then save the workbook as a PDF using C#. | Demonstrate placing default‑formatted rows at a specific worksheet position and exporting the result to PDF with Aspose.Cells in C#.
// Common Searches: C# Aspose.Cells insert blank rows at row 51 and generate PDF | how to place rows with default formatting in Excel using Aspose.Cells before PDF conversion | Aspose.Cells preserve formatting when adding rows then export workbook to PDF | convert edited Excel file to PDF after inserting rows with Aspose.Cells .NET
// Tags: Aspose.Cells insert rows C# | Aspose.Cells default formatting for new rows | Aspose.Cells export workbook to PDF | Aspose.Cells worksheet row insertion .NET

using System;
using Aspose.Cells;

namespace AsposeCellsInsertRowsAndPdf
{
    // Loads an existing Excel file, adds two blank rows at zero‑based index 50 on the first worksheet (keeping default formatting), and saves the modified workbook as a PDF.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (you can change the index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Insert two rows at row index 50 (zero‑based). 
            // The inserted rows will have default formatting (same as the rows above).
            worksheet.Cells.InsertRows(50, 2);

            // Save the modified workbook as a PDF. The format is inferred from the file extension.
            workbook.Save("output.pdf");

            Console.WriteLine("Rows inserted and PDF generated successfully.");
        }
    }
}

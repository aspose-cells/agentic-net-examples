// Title: C# – Delete Column Z, Hide Rows 50‑55 and Export Excel to PDF with Aspose.Cells
// Description: Loads an Excel workbook, removes column Z, hides rows 50 through 55, ensures the output directory exists, and saves the modified file as a PDF using Aspose.Cells for .NET. Includes basic error handling and automatic workbook creation if the source file is missing.
// Keywords: Aspose.Cells | C# | .NET | delete column Z | hide rows 50-55 | Excel to PDF conversion | workbook manipulation | column removal | row hiding | PDF export | automated reporting
// Common Searches: Aspose.Cells delete column Z C# | how to hide rows 50 to 55 with Aspose.Cells | export modified Excel to PDF using Aspose.Cells .NET | C# code to remove a column and hide rows before PDF conversion | Aspose.Cells example for column deletion and row hiding
// Developer Intent: Remove column Z, conceal rows 50‑55, and generate a PDF from the edited Excel workbook.
// Use Cases: Create a clean printable report by stripping an unnecessary column and hiding draft rows prior to PDF generation. | Prepare financial statements that exclude confidential data in column Z and suppress placeholder rows before distribution. | Automate batch processing of workbooks to enforce a standard layout (column Z removed, rows 50‑55 hidden) and produce PDF files for archiving.
// AI Prompts: Generate C# code with Aspose.Cells that deletes column Z, hides rows 50‑55, and saves the workbook as a PDF. | Explain step‑by‑step how to remove a specific column and hide a range of rows in an Excel file using Aspose.Cells before converting it to PDF. | Provide a robust Aspose.Cells .NET example that checks for the source file, creates missing directories, deletes column Z, hides rows 50‑55, and exports the result to PDF.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, removes column Z, hides rows 50 through 55, ensures the output directory exists, and saves the modified file as a PDF using Aspose.Cells for .NET. Includes basic error handling and automatic workbook creation if the source file is missing.
class Program
{
    static void Main()
    {
        // Paths for input Excel and output PDF
        string inputPath = @"C:\Input\sample.xlsx";
        string outputPath = @"C:\Output\result.pdf";

        try
        {
            // Ensure input file exists; create a simple workbook if missing
            if (!File.Exists(inputPath))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(inputPath));
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                ws.Cells["A1"].PutValue("Sample Data");
                wb.Save(inputPath);
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Delete column Z (zero‑based index 25)
            cells.DeleteColumn(25);

            // Hide rows 50 to 55 (zero‑based indices 49‑54, total 6 rows)
            cells.HideRows(49, 6);

            // Ensure output directory exists
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

            // Save the modified workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

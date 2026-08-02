// Title: C# – Remove Column Q from an Excel sheet and export it as PDF using Aspose.Cells
// Description: Load an existing workbook with Aspose.Cells for .NET, delete the Q column (index 16) from the first worksheet, and save the updated sheet directly to a PDF file.
// Keywords: Aspose.Cells delete column C# | remove column Q Excel | Excel to PDF conversion .NET | SaveFormat.Pdf Aspose | C# Excel column removal | Aspose.Cells PDF export | global Excel automation
// Common Searches: how to delete a column in Excel with Aspose.Cells | convert modified worksheet to PDF using Aspose | C# code to remove column Q and save as PDF | Aspose.Cells delete column and export PDF example
// Developer Intent: Programmatically eliminate column Q from an Excel workbook and generate a PDF of the cleaned worksheet.
// Use Cases: Redact sensitive data before publishing a spreadsheet as a PDF report. | Trim template files by discarding unused columns and creating printable PDFs. | Automate generation of distribution‑ready PDFs after column‑level cleanup.
// AI Prompts: Generate C# code that uses Aspose.Cells to delete column Q from a workbook and save the result as a PDF. | Create a reusable method that accepts input and output paths, removes a specified column index, and exports the worksheet to PDF. | Explain error‑handling strategies when deleting columns and converting Excel to PDF with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Load an existing workbook with Aspose.Cells for .NET, delete the Q column (index 16) from the first worksheet, and save the updated sheet directly to a PDF file.
class Program
{
    static void Main()
    {
        // Load the existing workbook from a file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet (you can change the index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Delete column Q (zero‑based index 16)
        worksheet.Cells.DeleteColumn(16);

        // Save the modified workbook as a PDF document
        string outputPath = "output.pdf";
        workbook.Save(outputPath, SaveFormat.Pdf);
    }
}

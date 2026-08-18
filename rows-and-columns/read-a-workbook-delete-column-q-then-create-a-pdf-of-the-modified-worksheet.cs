// Title: C# – Delete Column Q and Export Worksheet as PDF using Aspose.Cells
// Description: Load an Excel file with Aspose.Cells, remove column Q (index 16) from the first worksheet, and save the updated sheet directly as a PDF document.
// Keywords: Aspose.Cells delete column C# | remove column Q Excel | Excel to PDF conversion .NET | Aspose.Cells column removal | C# export worksheet PDF
// Common Searches: how to delete a specific column with Aspose.Cells | convert modified Excel sheet to PDF in C# | Aspose.Cells remove column by index | C# export worksheet after column deletion
// Developer Intent: Eliminate column Q from an Excel workbook and generate a PDF of the cleaned worksheet.
// Use Cases: Hide confidential data before publishing a report as PDF. | Automate template cleanup by stripping unwanted columns and creating distribution‑ready PDFs. | Process a batch of workbooks to remove a designated column and output each as a PDF file.
// AI Prompts: Generate C# code that uses Aspose.Cells to delete column Q (index 16) from the first sheet and save the result as a PDF. | Show an Aspose.Cells example that removes a column by name and exports the worksheet to PDF with default page settings. | Create a reusable method that accepts input and output paths, deletes column Q from the first worksheet, and returns the PDF as a byte array.

using System;
using Aspose.Cells;

// Load an Excel file with Aspose.Cells, remove column Q (index 16) from the first worksheet, and save the updated sheet directly as a PDF document.
class Program
{
    static void Main()
    {
        // Load the existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Delete column Q (zero‑based index 16)
        sheet.Cells.DeleteColumn(16);

        // Save the modified workbook as a PDF document
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}

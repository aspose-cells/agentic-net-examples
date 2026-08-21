// Title: C# – Delete rows 10‑15 in an Excel sheet and save as PDF with Aspose.Cells
// Description: Loads an Excel workbook, removes rows 10 through 15 from the first worksheet using Aspose.Cells, and directly saves the result as a PDF file.
// Keywords: Aspose.Cells | C# delete rows | remove rows Excel | Excel to PDF conversion | Aspose.Cells PDF export | Delete multiple rows | Workbook.Save PDF | Aspose.Cells .NET
// Common Searches: Aspose.Cells delete rows 10-15 C# | How to export modified Excel to PDF using Aspose.Cells | Remove specific rows from worksheet before PDF conversion | C# code to trim Excel rows and create PDF | Aspose.Cells delete rows and save as PDF example
// Developer Intent: Need to programmatically eliminate rows 10‑15 from an Excel worksheet and generate a PDF of the cleaned document.
// Use Cases: Preparing financial statements by stripping placeholder rows prior to distribution | Automating report cleanup in a server‑side .NET service | Generating printable PDFs from templates after removing temporary data rows | Batch processing of spreadsheets to delete header/footer rows before archiving as PDF
// AI Prompts: Provide C# Aspose.Cells code that deletes rows 10‑15 from the first sheet and exports the workbook to PDF. | Explain how to calculate the zero‑based start index for DeleteRows when the sheet contains hidden rows. | Show how to delete several non‑adjacent row ranges and then save the worksheet as a PDF using Aspose.Cells. | Demonstrate error handling for missing input file while performing row deletion and PDF conversion in C#.

using Aspose.Cells;

// Loads an Excel workbook, removes rows 10 through 15 from the first worksheet using Aspose.Cells, and directly saves the result as a PDF file.
class Program
{
    static void Main()
    {
        // Paths for the input Excel file and the output PDF file
        string inputPath = "input.xlsx";
        string outputPath = "output.pdf";

        // Load the existing workbook
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Delete rows 10 through 15 (zero‑based index starts at 9, total 6 rows)
        worksheet.Cells.DeleteRows(9, 6);

        // Save the modified workbook as a PDF document
        workbook.Save(outputPath, SaveFormat.Pdf);
    }
}

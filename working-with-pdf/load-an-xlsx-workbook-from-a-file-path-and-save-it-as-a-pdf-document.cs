// Title: Convert an XLSX workbook to PDF in C# using Aspose.Cells
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells and saves it as a PDF using SaveFormat.Pdf. | Show a concise Aspose.Cells example that reads an Excel workbook from a file path and exports it to a PDF document. | Provide a step‑by‑step C# snippet to convert a specific Excel file (input.xlsx) to output.pdf with Aspose.Cells.
// Common Searches: aspocells c# convert specific xlsx file to pdf | how to export Excel workbook to PDF using Aspose.Cells SaveFormat.Pdf | c# code example for loading workbook from path and saving as pdf | Aspose.Cells PDF export from XLSX in a console application | convert input.xlsx to output.pdf programmatically with Aspose.Cells
// Tags: Aspose.Cells convert XLSX to PDF C# | Load Excel workbook from file path Aspose.Cells | Save workbook as PDF using SaveFormat.Pdf | C# Excel to PDF conversion example | Export spreadsheet to PDF with Aspose.Cells

using Aspose.Cells;

// // Loads "input.xlsx" with Aspose.Cells and saves it as "output.pdf" by calling workbook.Save with SaveFormat.Pdf.
class Program
{
    static void Main()
    {
        // Load the XLSX workbook from the specified file path
        Workbook workbook = new Workbook("input.xlsx");

        // Save the loaded workbook as a PDF document
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}

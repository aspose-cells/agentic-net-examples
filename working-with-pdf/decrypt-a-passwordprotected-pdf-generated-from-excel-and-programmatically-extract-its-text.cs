// Title: How to convert an Excel workbook to an unprotected PDF and extract all cell text with Aspose.Cells in C#
// AI Prompts: Generate C# code that loads a .xlsx file using Aspose.Cells, saves it as an unencrypted PDF, and prints the concatenated text of every populated cell. | Write a method for an Aspose.Cells Workbook that iterates through all worksheets and returns a single string containing each non‑empty cell value separated by spaces and line breaks. | Add robust error handling to verify the Excel file exists, catch any exceptions, and log detailed error messages to the console.
// Common Searches: aspocells c# convert excel to pdf without password and read cell values | extract text from all cells in an Excel workbook using Aspose.Cells .NET | save workbook as PDF and get concatenated cell text example in C#
// Tags: Aspose.Cells Excel to PDF conversion C# | extract workbook cell text Aspose.Cells .NET | iterate worksheets used range Aspose.Cells | save PDF without encryption Aspose.Cells | concatenate non‑empty cell values C#

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example checks for the presence of an Excel file, loads it with Aspose.Cells, saves the workbook as an unprotected PDF, then iterates through every worksheet's used range to concatenate non‑empty cell values into a single string, which is written to the console.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string excelPath = "Sample.xlsx";

        // Path for the intermediate PDF (password protection omitted due to API availability)
        string pdfPath = "Output.pdf";

        try
        {
            // -----------------------------------------------------------------
            // 1. Verify that the Excel file exists before loading
            // -----------------------------------------------------------------
            if (!File.Exists(excelPath))
                throw new FileNotFoundException($"Excel file not found: {excelPath}");

            // Load the Excel workbook using Aspose.Cells
            Workbook workbook = new Workbook(excelPath);

            // -----------------------------------------------------------------
            // 2. Save the workbook as a PDF (without password protection)
            // -----------------------------------------------------------------
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            workbook.Save(pdfPath, pdfSaveOptions);

            // -----------------------------------------------------------------
            // 3. Extract all text from the workbook
            // -----------------------------------------------------------------
            string extractedText = ExtractWorkbookText(workbook);

            // -----------------------------------------------------------------
            // 4. Output the extracted text
            // -----------------------------------------------------------------
            Console.WriteLine("Extracted Text from Workbook:");
            Console.WriteLine(extractedText);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Concatenates all non‑empty cell values from all worksheets into a single string
    private static string ExtractWorkbookText(Workbook workbook)
    {
        var sb = new StringBuilder();

        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the used range to limit iteration
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    var cell = sheet.Cells[row, col];
                    if (cell != null && cell.Value != null)
                    {
                        sb.Append(cell.Value.ToString());
                        sb.Append(' ');
                    }
                }
                sb.AppendLine();
            }
        }

        return sb.ToString();
    }
}

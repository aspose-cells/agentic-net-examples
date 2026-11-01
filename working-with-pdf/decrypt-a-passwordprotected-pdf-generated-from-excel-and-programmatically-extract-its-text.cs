// Title: Decrypt a password‑protected Excel workbook, convert it to an unprotected PDF, and extract worksheet text using Aspose.Cells for .NET
// AI Prompts: Load an encrypted .xlsx file with a password using Aspose.Cells LoadOptions, remove workbook protection, and save the workbook as a PDF, returning the PDF file path. | Iterate through each worksheet of the decrypted workbook, read all used cells, and concatenate their values into a plain‑text string.
// Common Searches: C# Aspose.Cells load password protected Excel file and save as PDF | How to programmatically unprotect an Excel workbook with Aspose.Cells .NET | Extract all cell values from a decrypted Excel workbook using Aspose.Cells | Convert encrypted Excel to PDF and retrieve text content in C#
// Tags: load password protected workbook Aspose.Cells | unprotect Excel workbook .NET | save workbook as PDF Aspose.Cells | extract worksheet cell text C# | convert encrypted Excel to PDF

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example checks for a protected Excel file, loads it with the supplied password via LoadOptions, removes workbook protection, saves the workbook as an unprotected PDF, then walks through each worksheet's used range to collect cell values into a plain‑text string, which is printed to the console.
class ExcelDecryptAndExtract
{
    static void Main()
    {
        // Path to the password‑protected Excel file
        string excelPath = @"C:\Files\protected.xlsx";

        // Password used to protect the Excel file
        string excelPassword = "myPassword";

        // Verify that the source file exists
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Error: File not found - {excelPath}");
            return;
        }

        try
        {
            // Load the protected workbook with the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = excelPassword
            };
            Workbook workbook = new Workbook(excelPath, loadOptions);

            // Remove workbook protection (if any) using the same password
            workbook.Unprotect(excelPassword);

            // Save the workbook as an unprotected PDF
            string pdfPath = @"C:\Files\decrypted.pdf";
            workbook.Save(pdfPath, SaveFormat.Pdf);
            Console.WriteLine($"Decrypted PDF saved to: {pdfPath}");

            // Extract all text from the workbook (cell values)
            StringBuilder extractedText = new StringBuilder();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                extractedText.AppendLine($"--- Sheet: {sheet.Name} ---");
                Cells cells = sheet.Cells;

                // Iterate through used range only
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        var cell = cells[row, col];
                        if (cell != null && cell.Value != null)
                        {
                            extractedText.Append(cell.Value.ToString());
                        }
                        extractedText.Append('\t');
                    }
                    extractedText.AppendLine();
                }
                extractedText.AppendLine();
            }

            // Output the extracted text (or process as needed)
            Console.WriteLine("Extracted Text:");
            Console.WriteLine(extractedText.ToString());
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}

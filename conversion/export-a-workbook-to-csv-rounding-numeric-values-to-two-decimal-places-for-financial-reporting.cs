// Title: Export Workbook to CSV with Two‑Decimal Formatting using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, apply a "#,##0.00" style to numeric cells, ensure the output folder exists, and save the file as CSV so that financial values appear rounded to two decimal places.
// Keywords: Aspose.Cells CSV export | C# number format two decimal | financial report CSV | custom style Aspose.Cells | SaveFormat.Csv .NET | worksheet range formatting | directory creation C# | rounded values CSV
// Common Searches: Aspose.Cells export CSV with two decimal places | apply custom number format before saving CSV Aspose | C# round numbers to 2 decimals in CSV output
// Developer Intent: Generate a CSV file from an Aspose.Cells workbook where monetary figures are displayed with two‑decimal precision.
// Use Cases: Produce a financial summary that can be opened in Excel or any spreadsheet program with proper currency formatting. | Automate reporting pipelines that require CSV files with consistent decimal rounding. | Prevent file‑system errors by creating the destination directory programmatically before saving.
// AI Prompts: Write C# code that uses Aspose.Cells to export a workbook to CSV with numbers formatted to two decimal places. | Explain why applying a custom number format before calling Workbook.Save preserves rounding in the CSV output. | Show how to create a style, apply it to a column range, and ensure the output folder exists when saving a CSV file.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a workbook, apply a "#,##0.00" style to numeric cells, ensure the output folder exists, and save the file as CSV so that financial values appear rounded to two decimal places.
class ExportWorkbookToCsv
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add header row
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Amount");

            // Add sample financial data
            sheet.Cells["A2"].PutValue("Revenue");
            sheet.Cells["B2"].PutValue(12345.6789);
            sheet.Cells["A3"].PutValue("Expense");
            sheet.Cells["B3"].PutValue(9876.54321);

            // Create a style that formats numbers with two decimal places
            Style twoDecimalStyle = workbook.CreateStyle();
            twoDecimalStyle.Custom = "#,##0.00";

            // Apply the style only to the number format of column B
            StyleFlag flag = new StyleFlag { NumberFormat = true };
            sheet.Cells.CreateRange("B2:B3").ApplyStyle(twoDecimalStyle, flag);

            // Define output file path
            string outputPath = "FinancialReport.csv";

            // Ensure the directory exists (if a directory part is present)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the workbook to CSV; formatted values will be written
            workbook.Save(outputPath, SaveFormat.Csv);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

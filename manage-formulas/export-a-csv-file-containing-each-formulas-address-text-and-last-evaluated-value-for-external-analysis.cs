// Title: Export each cell's formula, address, and evaluated value to a CSV file using Aspose.Cells for .NET (C#)
// Common Searches: c# export all formulas from an Excel workbook to CSV using Aspose.Cells | how to get cell address, formula text and evaluated result with Aspose.Cells .NET | write Excel formula list with sheet name to CSV file in C# | Aspose.Cells iterate worksheets and extract formulas with values | escaping commas and quotes when saving Excel data to CSV in C#
// Tags: Aspose.Cells export formulas to CSV | extract cell address and evaluated value .NET | iterate worksheets for formula cells C# | CSV field escaping for Excel output | generate formula report using Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, walks through every worksheet and cell, captures each formula's address, text, and last evaluated value, and writes the data to a UTF‑8 CSV file. It creates the output directory if needed and safely escapes CSV fields containing commas, quotes, or line breaks.
class ExportFormulasToCsv
{
    static void Main(string[] args)
    {
        // Input Excel file path (change as needed)
        string excelFilePath = @"C:\Input\SampleWorkbook.xlsx";

        // Output CSV file path
        string csvFilePath = @"C:\Output\FormulasExport.csv";

        try
        {
            // Verify input file exists
            if (!File.Exists(excelFilePath))
            {
                Console.WriteLine($"Input file not found: {excelFilePath}");
                return;
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(csvFilePath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Load the workbook using Aspose.Cells
            Workbook workbook = new Workbook(excelFilePath);

            // Prepare a StreamWriter for the CSV output
            using (StreamWriter writer = new StreamWriter(csvFilePath, false, Encoding.UTF8))
            {
                // Write CSV header
                writer.WriteLine("Sheet,Address,Formula,LastEvaluatedValue");

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Iterate through all cells that contain formulas
                    foreach (Cell cell in cells)
                    {
                        if (!string.IsNullOrEmpty(cell.Formula))
                        {
                            string address = cell.Name;
                            string formula = cell.Formula;
                            string evaluatedValue = cell.Value?.ToString() ?? string.Empty;

                            // Escape CSV fields as needed
                            string sheetName = EscapeCsvField(sheet.Name);
                            string addressEsc = EscapeCsvField(address);
                            string formulaEsc = EscapeCsvField(formula);
                            string valueEsc = EscapeCsvField(evaluatedValue);

                            // Write the CSV line
                            writer.WriteLine($"{sheetName},{addressEsc},{formulaEsc},{valueEsc}");
                        }
                    }
                }
            }

            Console.WriteLine("Formula export completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to escape CSV fields containing commas, quotes, or newlines
    private static string EscapeCsvField(string field)
    {
        if (field == null)
            return string.Empty;

        bool mustEscape = field.Contains(",") || field.Contains("\"") || field.Contains("\r") || field.Contains("\n");
        if (mustEscape)
        {
            string escaped = field.Replace("\"", "\"\"");
            return $"\"{escaped}\"";
        }
        else
        {
            return field;
        }
    }
}

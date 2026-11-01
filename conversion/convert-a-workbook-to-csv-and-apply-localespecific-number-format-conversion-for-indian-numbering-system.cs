// Title: Export an Excel workbook to CSV with Indian numbering format using Aspose.Cells for .NET
// Common Searches: how to save an .xlsx as .csv with Indian number formatting using Aspose.Cells in C# | Aspose.Cells apply custom number format '#,##,##0' for Indian locale before CSV export | C# set workbook culture to en-IN and export to CSV with Aspose.Cells | convert Excel to CSV preserving Indian numbering system with Aspose.Cells .NET | apply Indian numbering style to all numeric cells in Aspose.Cells workbook
// Tags: Aspose.Cells export to CSV with Indian numbering | Workbook CultureInfo en-IN Aspose.Cells | Custom numeric style #,##,##0 Aspose.Cells | Apply Indian number format to all cells Aspose.Cells | .xlsx to .csv conversion Aspose.Cells .NET

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The program loads an input.xlsx file with Aspose.Cells, sets the workbook's CultureInfo to en‑IN, creates a custom style using the Indian numbering pattern "#,##,##0.##", applies this style to every numeric cell in the first worksheet, and saves the result as output.csv. It includes checks for file existence and basic exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.csv";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Apply Indian locale for number formatting
            workbook.Settings.CultureInfo = new CultureInfo("en-IN");

            // Define Indian numbering format (e.g., 1,00,000)
            Style indianStyle = workbook.CreateStyle();
            indianStyle.Custom = "#,##,##0.##";

            // Apply the format to all numeric cells in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

            foreach (Cell cell in usedRange)
            {
                if (cell.Type == CellValueType.IsNumeric)
                {
                    cell.SetStyle(indianStyle);
                }
            }

            // Save the workbook as CSV
            workbook.Save(outputPath, SaveFormat.Csv);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

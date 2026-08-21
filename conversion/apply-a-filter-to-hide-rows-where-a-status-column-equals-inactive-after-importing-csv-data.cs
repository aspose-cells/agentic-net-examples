// Title: C# – Hide rows with Status = 'Inactive' using Aspose.Cells AutoFilter after CSV import
// Description: Loads a CSV into an Aspose.Cells Workbook, determines the data range, applies an AutoFilter to the whole sheet, and uses a custom NotEqual filter on the Status column (column C) to hide rows marked as "Inactive" before saving the result as an XLSX file.
// Keywords: Aspose.Cells C# AutoFilter | hide rows by column value | filter CSV data Aspose.Cells | NotEqual custom filter | Excel hide inactive rows programmatically | load CSV to workbook Aspose | C# Excel data cleaning
// Common Searches: Aspose.Cells hide rows where column equals value | C# AutoFilter CSV Aspose.Cells example | filter out inactive records in Excel using code | apply NotEqual filter with Aspose.Cells | remove rows with status inactive after CSV import
// Developer Intent: Programmatically conceal rows whose Status column contains "Inactive" after importing a CSV file.
// Use Cases: Generate a clean employee roster by loading a CSV and automatically hiding former staff marked as Inactive. | Prepare a product inventory report that excludes discontinued items by filtering the Status field before export. | Create a reusable data‑pre‑processing routine that removes inactive records from any CSV source using Aspose.Cells.
// AI Prompts: Show how to extend the code to filter out multiple status values such as "Inactive" and "Closed". | Demonstrate applying a custom filter to a different column index after loading a CSV with Aspose.Cells. | Explain how to clear the AutoFilter at runtime to reveal all hidden rows again.

using System;
using System.IO;
using Aspose.Cells;

// Loads a CSV into an Aspose.Cells Workbook, determines the data range, applies an AutoFilter to the whole sheet, and uses a custom NotEqual filter on the Status column (column C) to hide rows marked as "Inactive" before saving the result as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "data.csv";
            const string outputPath = "filtered_output.xlsx";

            // Verify that the input CSV file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the CSV file (first row is assumed to contain column headers)
            var loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet where the CSV data resides
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range of the worksheet
            int maxColumn = sheet.Cells.MaxDataColumn; // zero‑based index of the last column with data
            int maxRow = sheet.Cells.MaxDataRow;       // zero‑based index of the last row with data

            // Ensure there is data to process
            if (maxColumn < 0 || maxRow < 0)
            {
                Console.WriteLine("The worksheet does not contain any data.");
                return;
            }

            // Build the address of the bottom‑right cell (e.g., "D10")
            string bottomRight = CellIndexToName(maxColumn, maxRow);

            // Apply AutoFilter to the whole data range (including the header row)
            sheet.AutoFilter.Range = $"A1:{bottomRight}";

            // Index of the Status column (zero‑based). Adjust if the column is elsewhere.
            int statusColumnIndex = 2; // Column C

            // Verify that the status column exists within the data range
            if (statusColumnIndex <= maxColumn)
            {
                // Hide rows where Status = "Inactive" by filtering for NOT equal to "Inactive"
                sheet.AutoFilter.Custom(statusColumnIndex, FilterOperatorType.NotEqual, "Inactive");
                // Apply the filter – rows not matching the criteria become hidden
                sheet.AutoFilter.Refresh();
            }
            else
            {
                Console.WriteLine($"Status column index {statusColumnIndex} is outside the data range.");
            }

            // Save the filtered workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Filtered workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper: converts zero‑based column/row indexes to an Excel cell address (e.g., 2,5 -> "C6")
    static string CellIndexToName(int columnIndex, int rowIndex)
    {
        string columnName = "";
        int dividend = columnIndex + 1;
        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }
        // Row index is zero‑based; add 1 for the Excel row number
        return $"{columnName}{rowIndex + 1}";
    }
}

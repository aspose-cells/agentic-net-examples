// Title: C# – Validate and Highlight yyyy-MM-dd Dates in Column E with Aspose.Cells
// Description: Creates or loads a workbook, scans column E, skips blanks, verifies each value against the exact "yyyy-MM-dd" pattern using DateTime.TryParseExact, marks non‑conforming cells with a red fill, reports mismatches, and saves the file.
// Keywords: Aspose.Cells | C# date validation | yyyy-MM-dd format | highlight invalid dates | Excel column E validation | DateTime.TryParseExact | red background style | data integrity | worksheet cell style | Aspose.Cells example
// Common Searches: Aspose.Cells validate date format C# | How to highlight cells with wrong date format using Aspose.Cells | Check yyyy-MM-dd column in Excel with C# | Mark invalid dates red Aspose.Cells | C# code to enforce date pattern in Excel column | Validate column E dates before saving workbook
// Developer Intent: Ensure every non‑empty cell in column E follows the yyyy-MM-dd pattern and visually flag any violations before saving the workbook.
// Use Cases: Cleanse imported spreadsheets where dates must follow a standard format. | Prevent downstream errors by rejecting workbooks containing malformed dates. | Provide visual cues (red fill) for users to correct date entries. | Generate logs of cells with invalid dates for audit trails. | Automate data‑quality checks in ETL pipelines using Aspose.Cells.
// AI Prompts: Generate C# code using Aspose.Cells that iterates column E, validates dates with "yyyy-MM-dd" format, and applies a red background to invalid cells. | Write a function that returns a list of addresses of cells in column E that do not match the required date pattern. | Explain how to customize the error style (font color, border, pattern) for invalid date cells in Aspose.Cells. | Show how to log validation results to a text file while processing a workbook with Aspose.Cells. | Provide a version of the code that throws an exception instead of highlighting when an invalid date is found.

using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDateValidation
{
    // Creates or loads a workbook, scans column E, skips blanks, verifies each value against the exact "yyyy-MM-dd" pattern using DateTime.TryParseExact, marks non‑conforming cells with a red fill, reports mismatches, and saves the file.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook(); // lifecycle: create
                Worksheet worksheet = workbook.Worksheets[0];

                // Sample data for demonstration – in real scenario the workbook would already contain data
                worksheet.Cells["E1"].PutValue("2023-01-15"); // valid
                worksheet.Cells["E2"].PutValue("15/01/2023"); // invalid format
                worksheet.Cells["E3"].PutValue("2023-12-31"); // valid
                worksheet.Cells["E4"].PutValue("");           // empty – ignored

                // Define the expected date format
                const string expectedFormat = "yyyy-MM-dd";

                // Get the used range to limit iteration (optional, here we just iterate a reasonable number of rows)
                int maxRow = worksheet.Cells.MaxDataRow;

                // Flag to indicate if any cell violates the format
                bool hasInvalidDate = false;

                // Iterate through all cells in column E (zero‑based column index 4)
                for (int row = 0; row <= maxRow; row++)
                {
                    Cell cell = worksheet.Cells[row, 4]; // column E

                    // Skip empty or blank cells
                    if (cell.Type == CellValueType.IsNull || string.IsNullOrWhiteSpace(cell.StringValue))
                        continue;

                    // Retrieve the cell's displayed string value
                    string cellText = cell.StringValue?.Trim() ?? string.Empty;

                    // Try to parse using the exact format
                    if (!DateTime.TryParseExact(
                            cellText,
                            expectedFormat,
                            CultureInfo.InvariantCulture,
                            DateTimeStyles.None,
                            out _))
                    {
                        // Mark the cell with a red background to highlight the issue
                        Style style = cell.GetStyle();
                        style.ForegroundColor = Color.Red;
                        style.Pattern = BackgroundType.Solid;
                        cell.SetStyle(style);

                        hasInvalidDate = true;
                    }
                }

                // Report any invalid dates
                if (hasInvalidDate)
                {
                    Console.WriteLine("One or more cells in column E do not match the format yyyy-MM-dd.");
                }

                // Ensure output directory exists
                string outputPath = "ValidatedWorkbook.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();

                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook (lifecycle: save)
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine("Workbook saved successfully.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

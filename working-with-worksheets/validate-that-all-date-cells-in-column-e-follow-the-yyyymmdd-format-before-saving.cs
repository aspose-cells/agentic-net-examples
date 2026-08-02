// Title: C# Aspose.Cells: Validate yyyy-MM-dd dates in column E before saving
// Description: Creates a workbook, fills column E with sample values, and iterates through all used rows. Each non‑empty cell is checked with DateTime.TryParseExact against the ISO 8601 pattern (yyyy‑MM‑dd). Invalid entries trigger an exception, while valid cells receive a custom date style before the workbook is saved.
// Keywords: Aspose.Cells | C# date validation | Excel date format yyyy-MM-dd | column E validation | Aspose.Cells .NET | custom date style | Excel workbook save | DateTime.TryParseExact
// Common Searches: Aspose.Cells validate column E date format | C# check yyyy-MM-dd in Excel cells | How to enforce date style with Aspose.Cells | Throw exception for invalid Excel dates Aspose | Iterate through Excel column with Aspose.Cells C#
// Developer Intent: Guarantee that every populated cell in column E contains a date string matching the yyyy‑MM‑dd format and apply that format before persisting the workbook.
// Use Cases: Screen user‑uploaded spreadsheets for incorrect date strings before further processing. | Standardize date columns in financial or inventory reports generated with Aspose.Cells. | Prevent downstream parsing errors by ensuring Excel dates follow the ISO 8601 standard.
// AI Prompts: Provide C# code that scans column E of an Aspose.Cells workbook, verifies each non‑empty cell matches the 'yyyy-MM-dd' pattern, and raises an exception for any mismatch. | Show how to set a custom date format on validated cells using Aspose.Cells styling API and then save the file. | Create NUnit tests that confirm the validation logic accepts correct dates and rejects invalid formats in an Aspose.Cells worksheet.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, fills column E with sample values, and iterates through all used rows. Each non‑empty cell is checked with DateTime.TryParseExact against the ISO 8601 pattern (yyyy‑MM‑dd). Invalid entries trigger an exception, while valid cells receive a custom date style before the workbook is saved.
public class ValidateDateColumnE
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data in column E (index 4)
            sheet.Cells["E1"].PutValue("2023-01-15");
            sheet.Cells["E2"].PutValue("2023/02/20"); // invalid format
            sheet.Cells["E3"].PutValue(DateTime.Now); // date value, will be formatted

            const string expectedFormat = "yyyy-MM-dd";

            // Determine the last used row in the worksheet
            int lastRow = sheet.Cells.MaxDataRow;

            // Validate each non‑empty cell in column E
            for (int row = 0; row <= lastRow; row++)
            {
                Cell cell = sheet.Cells[row, 4]; // column E (zero‑based index)

                if (cell.Type == CellValueType.IsNull) continue; // skip empty cells

                string cellText = cell.StringValue.Trim();

                // Verify the cell text matches the required date format
                if (!DateTime.TryParseExact(cellText, expectedFormat,
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out DateTime _))
                {
                    throw new InvalidOperationException(
                        $"Cell {cell.Name} does not follow the required format '{expectedFormat}'. Value: '{cellText}'.");
                }

                // Apply the correct display format to the cell
                Style style = cell.GetStyle();
                style.Custom = expectedFormat;
                cell.SetStyle(style);
            }

            // Save the workbook after successful validation
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ValidatedWorkbook.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred during validation: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        ValidateDateColumnE.Run();
    }
}

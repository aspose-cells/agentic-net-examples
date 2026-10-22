// Title: How to add exception handling for invalid ListObject TableCssId values in Aspose.Cells C#
// AI Prompts: Create a C# method that validates a ListObject TableCssId using a regular expression and throws an ArgumentException with a detailed message for invalid inputs. | Show how to wrap the TableCssId validation call in a try‑catch block that logs the exception message before saving the workbook. | Refactor the validation to use a custom InvalidTableCssIdException and update the example to catch this specific exception type.
// Common Searches: Aspose.Cells C# how to catch invalid TableCssId exception | validate ListObject identifier before saving workbook Aspose.Cells | C# regex pattern for TableCssId naming rules Aspose.Cells | exception handling example for ListObject display name in Aspose.Cells | what error is thrown for bad TableCssId in Aspose.Cells C#
// Tags: Aspose.Cells ListObject identifier validation | C# regex TableCssId pattern | exception handling invalid TableCssId Aspose.Cells | try-catch ListObject display name Aspose.Cells | custom InvalidTableCssIdException C#

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a workbook, adds a ListObject table, validates a custom TableCssId with a regular expression, throws a descriptive ArgumentException for invalid identifiers, and demonstrates try‑catch blocks for both identifier validation and workbook saving.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.2);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.8);

            // Define the range for the table (including header row)
            int firstRow = 0;      // zero‑based index
            int firstColumn = 0;
            int totalRows = 3;     // header + 2 data rows
            int totalColumns = 2;

            // Add a ListObject (table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Set Table identifier with validation
            try
            {
                SetTableIdentifier(table, "my-table-css"); // valid example
                // SetTableIdentifier(table, "123-invalid"); // uncomment to test invalid case
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error setting table identifier: {ex.Message}");
            }

            // Save the workbook
            string outputPath = "Result.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    /// <param name="table">The Aspose.Cells ListObject representing the table.</param>
    /// <param name="identifier">The identifier to assign.</param>
    static void SetTableIdentifier(ListObject table, string identifier)
    {
        // Validation rules:
        // 1. Must not be null, empty, or whitespace.
        // 2. Must start with a letter.
        // 3. May contain letters, digits, hyphens, or underscores only.
        if (string.IsNullOrWhiteSpace(identifier))
        {
            throw new ArgumentException("Identifier cannot be null, empty, or consist only of whitespace.");
        }

        if (!Regex.IsMatch(identifier, @"^[A-Za-z][A-Za-z0-9\-_]*$"))
        {
            throw new ArgumentException(
                $"Identifier '{identifier}' is invalid. It must start with a letter and contain only letters, digits, hyphens, or underscores.");
        }

        // Assign the validated identifier to the table's display name (as a stand‑in for TableCssId)
        table.DisplayName = identifier;
    }
}

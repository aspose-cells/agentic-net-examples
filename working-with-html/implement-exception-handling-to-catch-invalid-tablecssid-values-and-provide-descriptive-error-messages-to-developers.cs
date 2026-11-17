// Title: How to add validation and exception handling when assigning a TableStyleName (TableCssId) to an Excel table using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells validate table style name before setting TableStyleName | how to catch invalid TableCssId exception with Aspose.Cells | example of regex validation for Excel table identifier in Aspose.Cells | Aspose.Cells handling missing worksheet or table errors in C# | validate and assign table CSS identifier using Aspose.Cells API
// Tags: Aspose.Cells table style validation | C# Aspose.Cells ListObject exception handling | regex Excel table identifier validation | validate TableStyleName before assignment | missing worksheet error handling Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample demonstrates loading an Excel workbook with Aspose.Cells, checking that the target worksheet and table exist, validating a CSS identifier using a regular expression, assigning it to the table's TableStyleName, and saving the file. It includes specific catches for file-not-found, argument, and general exceptions, outputting clear error messages for developers.
public class TableCssIdHandler
{
    /// <param name="filePath">Full path to the Excel file.</param>
    /// <param name="sheetName">Name of the worksheet containing the table.</param>
    /// <param name="tableName">Name of the table (ListObject) to modify.</param>
    /// <param name="cssId">The identifier to assign (used here as TableStyleName).</param>
    public void ApplyTableCssId(string filePath, string sheetName, string tableName, string cssId)
    {
        try
        {
            // Ensure the file exists before attempting to load it
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"The Excel file '{filePath}' was not found.");

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Retrieve the target worksheet
            Worksheet sheet = workbook.Worksheets[sheetName];
            if (sheet == null)
                throw new ArgumentException($"Worksheet '{sheetName}' does not exist in the workbook.");

            // Retrieve the target table (ListObject)
            ListObject table = sheet.ListObjects[tableName];
            if (table == null)
                throw new ArgumentException($"Table '{tableName}' does not exist on worksheet '{sheetName}'.");

            // Validate the identifier before assignment
            ValidateTableCssId(cssId);

            // Assign the identifier.
            // Aspose.Cells does not expose a TableCssId property; using TableStyleName as a demonstrative placeholder.
            table.TableStyleName = cssId;

            // Save the workbook (overwrites the original file)
            workbook.Save(filePath);
        }
        catch (ArgumentException ex)
        {
            // Specific handling for validation errors
            Console.Error.WriteLine($"Validation error: {ex.Message}");
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine($"File error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // General error handling for any other unexpected issues
            Console.Error.WriteLine($"Unexpected error while applying identifier: {ex.Message}");
        }
    }

    /// <param name="cssId">The identifier to validate.</param>
    private void ValidateTableCssId(string cssId)
    {
        if (string.IsNullOrWhiteSpace(cssId))
        {
            throw new ArgumentException("Identifier cannot be null, empty, or consist solely of whitespace.");
        }

        // Must start with a letter and may contain letters, digits, hyphens, or underscores.
        if (!Regex.IsMatch(cssId, @"^[A-Za-z][A-Za-z0-9_-]*$"))
        {
            throw new ArgumentException(
                $"Identifier '{cssId}' is invalid. It must start with a letter and contain only letters, digits, hyphens, or underscores.");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Example usage – adjust the parameters as needed.
        string filePath = "sample.xlsx";
        string sheetName = "Sheet1";
        string tableName = "Table1";
        string cssId = "MyTableStyle";

        // Verify that the input file exists before proceeding.
        if (!File.Exists(filePath))
        {
            Console.Error.WriteLine($"Input file not found: {filePath}");
            return;
        }

        try
        {
            TableCssIdHandler handler = new TableCssIdHandler();
            handler.ApplyTableCssId(filePath, sheetName, tableName, cssId);
            Console.WriteLine("Identifier applied successfully.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error during execution: {ex.Message}");
        }
    }
}

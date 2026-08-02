// Title: C# – Validate HtmlSaveOptions.TableCssId and Handle Exceptions in Aspose.Cells
// Description: Demonstrates how to verify a user‑provided TableCssId with a regular expression, throw a descriptive ArgumentException for invalid values, and wrap the HTML export of a workbook in try/catch blocks. The example shows workbook creation, CSS‑id validation, assignment to HtmlSaveOptions, and graceful error reporting for Aspose.Cells for .NET.
// Keywords: Aspose.Cells | HtmlSaveOptions | TableCssId validation | C# regex CSS identifier | ArgumentException handling | HTML export workbook | custom table CSS id | .NET | exception handling Aspose.Cells | validate CSS id Aspose
// Common Searches: how to validate TableCssId Aspose.Cells | TableCssId regex pattern C# | exception thrown for invalid TableCssId | catch ArgumentException when saving HTML with Aspose.Cells | custom CSS id for HTML export Aspose.Cells
// Developer Intent: Ensure the TableCssId meets CSS naming rules, throw a clear error if it does not, and prevent runtime failures during HTML conversion.
// Use Cases: Validate user input before assigning it to HtmlSaveOptions.TableCssId to avoid export errors. | Log or display a precise validation message when the CSS identifier is malformed. | Wrap workbook‑to‑HTML conversion in robust try/catch logic to handle both validation failures and unexpected runtime issues.
// AI Prompts: Write a C# method that checks HtmlSaveOptions.TableCssId against CSS naming rules and raises a detailed ArgumentException. | Show how to modify the try‑catch block to write TableCssId validation errors to a log file instead of the console. | Create unit tests for ValidateTableCssId covering valid IDs, empty strings, whitespace, and illegal characters.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to verify a user‑provided TableCssId with a regular expression, throw a descriptive ArgumentException for invalid values, and wrap the HTML export of a workbook in try/catch blocks. The example shows workbook creation, CSS‑id validation, assignment to HtmlSaveOptions, and graceful error reporting for Aspose.Cells for .NET.
    public class HtmlSaveOptionsTableCssIdValidationDemo
    {
        // Validates that the TableCssId follows a simple CSS identifier pattern.
        private static void ValidateTableCssId(string cssId)
        {
            // CSS identifiers must start with a letter and can contain letters, digits, hyphens, and underscores.
            if (string.IsNullOrWhiteSpace(cssId))
                throw new ArgumentException("TableCssId cannot be null, empty, or whitespace.");

            // Simple regex for validation.
            if (!Regex.IsMatch(cssId, @"^[a-zA-Z][\w-]*$"))
                throw new ArgumentException($"TableCssId \"{cssId}\" is invalid. It must start with a letter and contain only letters, digits, hyphens, or underscores.");
        }

        public static void Run()
        {
            try
            {
                // Create a sample workbook with test data.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                worksheet.Cells["A1"].PutValue("Name");
                worksheet.Cells["B1"].PutValue("Age");
                worksheet.Cells["A2"].PutValue("John");
                worksheet.Cells["B2"].PutValue(30);
                worksheet.Cells["A3"].PutValue("Alice");
                worksheet.Cells["B3"].PutValue(25);

                // Configure HTML save options.
                HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

                // Example of a valid TableCssId.
                string userProvidedCssId = "custom-table-style";

                // Validate before assigning.
                ValidateTableCssId(userProvidedCssId);
                saveOptions.TableCssId = userProvidedCssId;

                // Save the workbook with the specified HTML options.
                workbook.Save("output.html", saveOptions);

                Console.WriteLine($"HTML file saved successfully with TableCssId: \"{saveOptions.TableCssId}\"");
            }
            catch (ArgumentException ex)
            {
                // Provide a clear, descriptive error message for developers.
                Console.Error.WriteLine($"TableCssId validation error: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Catch any other unexpected exceptions.
                Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application.
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                HtmlSaveOptionsTableCssIdValidationDemo.Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }
}

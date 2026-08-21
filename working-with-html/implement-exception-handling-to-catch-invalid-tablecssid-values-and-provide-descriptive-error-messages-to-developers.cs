// Title: C# – Validate HtmlSaveOptions.TableCssId with Exception Handling in Aspose.Cells
// Description: Demonstrates how to verify a custom TableCssId before assigning it to HtmlSaveOptions, catch ArgumentException for invalid identifiers, and handle unexpected errors while exporting a workbook to HTML using Aspose.Cells for .NET.
// Keywords: Aspose.Cells HtmlSaveOptions TableCssId | C# TableCssId validation | exception handling Aspose.Cells | ArgumentException HtmlSaveOptions | custom CSS ID export
// Common Searches: validate TableCssId Aspose.Cells C# | HtmlSaveOptions TableCssId exception example | how to catch invalid TableCssId in Aspose.Cells | C# Aspose.Cells HTML export CSS id validation
// Developer Intent: Ensure only valid TableCssId values are set on HtmlSaveOptions and provide clear error messages when validation fails.
// Use Cases: Prevent runtime failures when exporting large reports to HTML. | Log precise validation errors in CI/CD pipelines. | Enforce corporate naming conventions for CSS IDs during automated workbook conversion.
// AI Prompts: Generate a ValidateTableCssId method that permits only alphanumeric characters, hyphens, and underscores. | Create a custom TableCssIdValidationException and replace ArgumentException in the sample. | Write NUnit tests covering null, empty, whitespace, and illegal characters for ValidateTableCssId.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to verify a custom TableCssId before assigning it to HtmlSaveOptions, catch ArgumentException for invalid identifiers, and handle unexpected errors while exporting a workbook to HTML using Aspose.Cells for .NET.
    public class HtmlSaveOptionsTableCssIdValidationDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and add some sample data
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Cells["A1"].PutValue("Name");
                worksheet.Cells["B1"].PutValue("Age");
                worksheet.Cells["A2"].PutValue("John");
                worksheet.Cells["B2"].PutValue(30);

                // Prepare HTML save options
                HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

                // Desired TableCssId value
                string desiredTableCssId = "custom-table-style";

                // Validate the TableCssId before assigning it
                ValidateTableCssId(desiredTableCssId);

                // Assign the validated value
                saveOptions.TableCssId = desiredTableCssId;

                // Save the workbook using the configured options
                workbook.Save("output.html", saveOptions);

                Console.WriteLine($"Workbook saved successfully with TableCssId: '{saveOptions.TableCssId}'.");
            }
            catch (ArgumentException ex)
            {
                // Handle validation errors for TableCssId
                Console.Error.WriteLine($"TableCssId validation error: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Handle any other unexpected errors
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Simple validation logic for TableCssId
        private static void ValidateTableCssId(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
                throw new ArgumentException("TableCssId cannot be null, empty, or consist only of whitespace.");

            // Disallow whitespace characters within the identifier
            if (id.IndexOfAny(new char[] { ' ', '\t', '\r', '\n' }) >= 0)
                throw new ArgumentException("TableCssId must not contain whitespace characters.");

            // Additional custom validation rules can be added here if needed
        }
    }
}

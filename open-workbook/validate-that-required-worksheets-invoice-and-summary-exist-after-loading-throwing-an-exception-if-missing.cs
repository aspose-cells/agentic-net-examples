// Title: C# – Validate presence of "Invoice" and "Summary" worksheets in an Excel file with Aspose.Cells
// Description: Loads an Excel workbook using Aspose.Cells for .NET, scans the Worksheets collection case‑insensitively for "Invoice" and "Summary", and throws an InvalidOperationException with a clear message if either sheet is missing; otherwise writes a success message to the console.
// Keywords: Aspose.Cells | C# worksheet validation | Excel workbook check | required sheets Invoice Summary | InvalidOperationException | .NET Excel validation | load workbook Aspose | case insensitive sheet name | global developers
// Common Searches: Aspose.Cells check if worksheet exists | C# verify Invoice sheet in Excel | throw exception missing Summary sheet Aspose | validate required worksheets .NET | how to ensure Excel template contains specific sheets
// Developer Intent: Confirm that an Excel workbook contains both the "Invoice" and "Summary" sheets before processing, and halt execution with a clear exception when a required sheet is absent.
// Use Cases: Pre‑flight validation of an invoice‑processing template to guarantee required sheets are present before data extraction. | Automated reporting pipeline that aborts with a descriptive error when the summary worksheet is missing from incoming Excel files. | Desktop application that loads user‑provided workbooks and must verify required sheets exist to prevent runtime failures.
// AI Prompts: Generate a reusable method that accepts a list of required worksheet names and throws an InvalidOperationException for any that are missing using Aspose.Cells. | Refactor the validation logic to return a collection of missing worksheet names instead of throwing exceptions, and show how to handle the result. | Write unit tests for ValidateRequiredWorksheets covering scenarios where both sheets exist, only one exists, and neither exists.

using System;
using Aspose.Cells;

namespace WorksheetValidationExample
{
    // Loads an Excel workbook using Aspose.Cells for .NET, scans the Worksheets collection case‑insensitively for "Invoice" and "Summary", and throws an InvalidOperationException with a clear message if either sheet is missing; otherwise writes a success message to the console.
    public class Validator
    {
        /// <param name="filePath">Path to the Excel file to validate.</param>
        public static void ValidateRequiredWorksheets(string filePath)
        {
            // Load the workbook using the standard constructor (lifecycle rule: load)
            Workbook workbook = new Workbook(filePath);

            // Check for the presence of the "Invoice" worksheet
            bool hasInvoice = false;
            // Check for the presence of the "Summary" worksheet
            bool hasSummary = false;

            // Iterate through the worksheets collection
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name.Equals("Invoice", StringComparison.OrdinalIgnoreCase))
                {
                    hasInvoice = true;
                }
                else if (sheet.Name.Equals("Summary", StringComparison.OrdinalIgnoreCase))
                {
                    hasSummary = true;
                }

                // If both are found, no need to continue looping
                if (hasInvoice && hasSummary)
                {
                    break;
                }
            }

            // Throw detailed exceptions if required sheets are missing
            if (!hasInvoice)
            {
                throw new InvalidOperationException("Required worksheet \"Invoice\" is missing.");
            }

            if (!hasSummary)
            {
                throw new InvalidOperationException("Required worksheet \"Summary\" is missing.");
            }

            // If execution reaches here, both worksheets are present
            Console.WriteLine("Validation successful: both \"Invoice\" and \"Summary\" worksheets are present.");
        }

        // Example usage
        public static void Main()
        {
            try
            {
                // Replace with the actual path to your Excel file
                string excelPath = "SampleWorkbook.xlsx";
                ValidateRequiredWorksheets(excelPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Validation failed: {ex.Message}");
            }
        }
    }
}

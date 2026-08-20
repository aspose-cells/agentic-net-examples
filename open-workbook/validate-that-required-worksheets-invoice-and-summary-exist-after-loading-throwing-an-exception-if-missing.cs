// Title: C# – Validate that 'Invoice' and 'Summary' worksheets exist in an Excel file using Aspose.Cells
// Description: Loads an Excel workbook with Aspose.Cells for .NET, checks case‑insensitive presence of the required 'Invoice' and 'Summary' sheets, throws an InvalidOperationException if a sheet is missing, and writes a success message when both are found.
// Keywords: Aspose.Cells | C# workbook validation | check worksheet existence | required sheets | Invoice sheet | Summary sheet | throw exception | Excel template validation | load workbook | pre‑flight check
// Common Searches: Aspose.Cells verify worksheet exists | C# check if Excel sheet named Invoice is present | Throw error when required sheet missing Aspose.Cells | Validate Excel template before processing with Aspose | How to ensure Summary sheet exists in workbook using Aspose.Cells
// Developer Intent: Confirm that a loaded Excel workbook contains both the 'Invoice' and 'Summary' worksheets and stop execution with a clear exception if either is absent.
// Use Cases: Automated invoice processing pipelines that require an 'Invoice' sheet | Reporting jobs that depend on a 'Summary' worksheet | Pre‑flight validation of user‑uploaded Excel templates | Quality control for data import routines in ERP systems | CI/CD tests that verify Excel file structure
// AI Prompts: Write a reusable method that accepts a list of required worksheet names and returns missing ones in an Aspose.Cells workbook. | Convert the validation loop to a LINQ expression for cleaner C# code. | Create NUnit tests that mock a Workbook with and without the required sheets to verify exception handling. | Generate a PowerShell script that calls the C# validator for batch Excel files. | Explain how to log validation results to a file instead of console using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsValidationExample
{
    // Loads an Excel workbook with Aspose.Cells for .NET, checks case‑insensitive presence of the required 'Invoice' and 'Summary' sheets, throws an InvalidOperationException if a sheet is missing, and writes a success message when both are found.
    public class WorkbookValidator
    {
        /// <param name="filePath">Path to the Excel file to validate.</param>
        public static void ValidateRequiredWorksheets(string filePath)
        {
            // Load the workbook using the standard constructor (lifecycle rule)
            Workbook workbook = new Workbook(filePath);

            // Check for the presence of the "Invoice" worksheet
            bool hasInvoice = false;
            // Check for the presence of the "Summary" worksheet
            bool hasSummary = false;

            // Iterate through the worksheet collection
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
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

            // Throw detailed exceptions if any required worksheet is missing
            if (!hasInvoice)
            {
                throw new InvalidOperationException("Required worksheet 'Invoice' is missing in the workbook.");
            }

            if (!hasSummary)
            {
                throw new InvalidOperationException("Required worksheet 'Summary' is missing in the workbook.");
            }

            // Optional: indicate successful validation
            Console.WriteLine("Validation succeeded: both 'Invoice' and 'Summary' worksheets are present.");
        }

        // Example usage
        public static void Main(string[] args)
        {
            // Replace with the actual path to your Excel file
            string excelFilePath = "input.xlsx";

            try
            {
                ValidateRequiredWorksheets(excelFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Validation error: {ex.Message}");
            }
        }
    }
}

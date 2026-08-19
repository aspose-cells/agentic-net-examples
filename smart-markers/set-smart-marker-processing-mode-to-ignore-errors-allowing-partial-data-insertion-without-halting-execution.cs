// Title: C# – Aspose.Cells Smart Marker Processing with Error Ignoring
// Description: Loads a workbook template containing smart markers, assigns a data source with missing fields, sets CalculationOptions.IgnoreError to true, and calls WorkbookDesigner.Process(true) to continue processing and preserve unrecognized markers, then saves the result.
// Keywords: Aspose.Cells | C# | .NET | smart markers | ignore errors | CalculationOptions.IgnoreError | WorkbookDesigner.Process | partial data insertion | error handling | report generation
// Common Searches: Aspose.Cells ignore smart marker errors C# | continue processing smart markers when data is missing | WorkbookDesigner.Process preserve unknown markers | CalculationOptions.IgnoreError example | smart marker partial insertion .NET
// Developer Intent: Configure smart marker processing to ignore data‑related errors so the operation completes without throwing exceptions.
// Use Cases: Generate a report from a template when some rows lack certain fields, using IgnoreError to produce a partial output. | Preserve custom or future smart markers in a workbook while skipping rows that cause insertion failures. | Automate batch report creation where data quality varies, ensuring the process never aborts due to missing values.
// AI Prompts: Show C# code that sets Aspose.Cells smart marker processing to ignore errors and continue. | Explain how CalculationOptions.IgnoreError works with WorkbookDesigner.Process(true) for missing fields. | Give an example of preserving unrecognized smart markers while ignoring data errors in Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerIgnoreErrorsDemo
{
    // Demonstrates how to configure smart marker processing to ignore errors,
    // allowing the operation to continue even when some data cannot be inserted.
    // Loads a workbook template containing smart markers, assigns a data source with missing fields, sets CalculationOptions.IgnoreError to true, and calls WorkbookDesigner.Process(true) to continue processing and preserve unrecognized markers, then saves the result.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Path to the template workbook containing smart markers.
                const string templatePath = "TemplateWithSmartMarkers.xlsx";

                // Verify that the template file exists to avoid FileNotFoundException.
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the workbook.
                Workbook workbook = new Workbook(templatePath);

                // Create a WorkbookDesigner and associate it with the loaded workbook.
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook,
                    // LineByLine is obsolete; retained for compatibility with older templates.
                    LineByLine = false
                };

                // Prepare a data source that intentionally has missing fields to trigger errors.
                var employees = new List<dynamic>
                {
                    new { Name = "John Doe", Age = 30 },               // Missing Salary
                    new { Name = "Jane Smith", Age = 28, Salary = 75000 } // Complete row
                };

                // Set the data source for the smart markers.
                designer.SetDataSource("Employees", employees);

                // Configure calculation options to ignore errors during formula evaluation.
                // This ensures that missing data or formula problems do not halt execution.
                var calcOptions = new CalculationOptions { IgnoreError = true };
                workbook.CalculateFormula(calcOptions);

                // Process the smart markers.
                // The boolean parameter 'true' tells the designer to preserve any unrecognized
                // smart markers, which also helps avoid exceptions for missing data.
                designer.Process(true);

                // Save the resulting workbook.
                const string outputPath = "SmartMarkersProcessed_IgnoringErrors.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

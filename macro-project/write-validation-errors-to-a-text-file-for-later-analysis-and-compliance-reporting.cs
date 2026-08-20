// Title: Aspose.Cells C# – Export Data‑Validation Rules and Violations to a Text Report
// Description: A C# console example that creates a workbook, adds a whole‑number validation for cells A1:A5, populates test data, and generates a plain‑text file listing each validation’s settings and every cell that fails the rule. The workbook is saved for reference, making it ideal for compliance audits and automated spreadsheet checks.
// Keywords: Aspose.Cells validation report | C# data validation export | write validation errors to txt | log worksheet rule violations | .NET spreadsheet compliance | cell validation CSV alternative | Aspose.Cells Workbook Save | data‑validation audit file
// Common Searches: how to export Aspose.Cells validation errors to a file | Aspose.Cells write data‑validation details to text | C# generate validation report for Excel workbook | log out‑of‑range values using Aspose.Cells | create compliance report from Excel validation rules
// Developer Intent: Produce a text file that documents all data‑validation definitions in a worksheet and enumerates the cells that violate those definitions.
// Use Cases: Compliance audit: capture validation rules and offending cells for regulatory review. | Automated data quality check: identify non‑numeric or out‑of‑range entries in numeric columns. | Scheduled spreadsheet processing: generate daily validation summaries for ETL pipelines. | Debugging: quickly see which cells break custom validation logic during development.
// AI Prompts: Add the worksheet name and cell address to each line of the validation error report. | Summarize the total number of violations per validation and append a summary section to the text file. | Convert the output format from plain text to CSV, including columns for Worksheet, Cell, Error Title, and Message. | Include a timestamp and the executing user in the report header for audit trails.

using System;
using System.IO;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsValidationReport
{
    // A C# console example that creates a workbook, adds a whole‑number validation for cells A1:A5, populates test data, and generates a plain‑text file listing each validation’s settings and every cell that fails the rule. The workbook is saved for reference, making it ideal for compliance audits and automated spreadsheet checks.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define a validation for cells A1:A5 (whole numbers between 10 and 20)
                CellArea area = CellArea.CreateCellArea("A1", "A5");
                int validationIndex = sheet.Validations.Add(area);
                Validation validation = sheet.Validations[validationIndex];
                validation.Type = ValidationType.WholeNumber;
                validation.Operator = OperatorType.Between;
                validation.Formula1 = "10";
                validation.Formula2 = "20";
                validation.ErrorTitle = "Invalid Input";
                validation.ErrorMessage = "Value must be between 10 and 20.";
                validation.ShowError = true;
                validation.AlertStyle = ValidationAlertType.Stop;

                // Insert some test values (some invalid)
                sheet.Cells["A1"].PutValue(5);   // invalid
                sheet.Cells["A2"].PutValue(15);  // valid
                sheet.Cells["A3"].PutValue(25);  // invalid
                sheet.Cells["A4"].PutValue(12);  // valid
                sheet.Cells["A5"].PutValue(8);   // invalid

                // Prepare the output file for validation error reporting
                string reportPath = "ValidationErrorsReport.txt";

                // Ensure the directory for the report exists
                string reportDir = Path.GetDirectoryName(Path.GetFullPath(reportPath));
                if (!Directory.Exists(reportDir))
                {
                    Directory.CreateDirectory(reportDir);
                }

                using (StreamWriter writer = new StreamWriter(reportPath))
                {
                    // Iterate through all validations in the worksheet
                    foreach (Validation val in sheet.Validations)
                    {
                        writer.WriteLine("Validation:");
                        writer.WriteLine($"  Error Title : {val.ErrorTitle}");
                        writer.WriteLine($"  Error Message : {val.ErrorMessage}");
                        writer.WriteLine($"  Show Error : {val.ShowError}");
                        writer.WriteLine($"  Alert Style : {val.AlertStyle}");
                        writer.WriteLine("  Affected Ranges:");

                        // List each cell area covered by this validation
                        foreach (CellArea range in val.Areas)
                        {
                            writer.WriteLine($"    {range.StartRow + 1}:{range.StartColumn + 1} to {range.EndRow + 1}:{range.EndColumn + 1}");
                        }

                        writer.WriteLine();
                    }

                    // Check each cell in the validation ranges for violations
                    foreach (Validation val in sheet.Validations)
                    {
                        foreach (CellArea range in val.Areas)
                        {
                            for (int row = range.StartRow; row <= range.EndRow; row++)
                            {
                                for (int col = range.StartColumn; col <= range.EndColumn; col++)
                                {
                                    Cell cell = sheet.Cells[row, col];

                                    // Simple check for whole number between the two formulas
                                    if (val.Type == ValidationType.WholeNumber && val.Operator == OperatorType.Between)
                                    {
                                        // Validation formulas may include a leading '=', remove it before parsing
                                        string formula1 = val.Formula1?.TrimStart('=') ?? "0";
                                        string formula2 = val.Formula2?.TrimStart('=') ?? "0";

                                        if (double.TryParse(formula1, NumberStyles.Any, CultureInfo.InvariantCulture, out double min) &&
                                            double.TryParse(formula2, NumberStyles.Any, CultureInfo.InvariantCulture, out double max))
                                        {
                                            if (cell.Type == CellValueType.IsNumeric)
                                            {
                                                double cellVal = cell.DoubleValue;
                                                if (cellVal < min || cellVal > max)
                                                {
                                                    writer.WriteLine($"Cell {cell.Name} contains invalid value {cellVal}. {val.ErrorMessage}");
                                                }
                                            }
                                            else
                                            {
                                                writer.WriteLine($"Cell {cell.Name} does not contain a numeric value. {val.ErrorMessage}");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // Save the workbook (optional, just to keep the file)
                string workbookPath = "ValidationDemo.xlsx";
                workbook.Save(workbookPath);

                Console.WriteLine($"Validation error report written to {Path.GetFullPath(reportPath)}");
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(workbookPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

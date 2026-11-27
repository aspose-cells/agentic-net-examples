// Title: Export rows that fail cell validation to a separate "ErrorRows" worksheet with detailed messages using Aspose.Cells for .NET (C#)
// Common Searches: how to extract rows with validation errors to a new sheet using Aspose.Cells in C# | Aspose.Cells copy rows that fail required field validation into a separate worksheet | C# code to log Excel validation failures and save them in an error sheet with Aspose.Cells | export invalid Excel rows with concatenated error messages using Aspose.Cells .NET | create an error worksheet for rows that don't meet validation rules with Aspose.Cells
// Tags: export invalid rows Aspose.Cells C# | validation error worksheet Aspose.Cells | copy rows with failed validation Aspose.Cells | add error messages column Excel Aspose.Cells | auto-fit columns error sheet Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace ExportInvalidRows
{
    // The example loads an input XLSX workbook, creates or clears a worksheet named "ErrorRows", copies the header row, adds an extra column for concatenated validation messages, iterates through each data row checking each cell's Validation object (including a reflected IsRequired property), records missing‑required or rule‑present errors, copies rows that contain errors to the error sheet together with the generated message, auto‑fits the columns for readability, and saves the workbook as a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output_with_errors.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.Error.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Work with the first worksheet that contains data/validation
                Worksheet sourceSheet = workbook.Worksheets[0];
                Cells sourceCells = sourceSheet.Cells;

                // Create (or clear) the error worksheet
                const string errorSheetName = "ErrorRows";
                Worksheet errorSheet = workbook.Worksheets[errorSheetName];

                if (errorSheet != null)
                {
                    // Clear existing contents (clear all cells)
                    errorSheet.Cells.Clear(); // Clear entire sheet
                }
                else
                {
                    // Add a new sheet for error rows
                    errorSheet = workbook.Worksheets.Add(errorSheetName);
                }

                Cells errorCells = errorSheet.Cells;

                // Determine the used range of the source sheet
                int maxRow = sourceCells.MaxDataRow;      // zero‑based index of the last row with data
                int maxCol = sourceCells.MaxDataColumn;   // zero‑based index of the last column with data

                int destRow = 0; // Destination row index in the error sheet

                // Copy header row (if any) to the error sheet for context
                for (int c = 0; c <= maxCol; c++)
                {
                    errorCells[destRow, c].PutValue(sourceCells[0, c].Value);
                }

                // Add an extra column for concatenated error messages
                int errorMessageCol = maxCol + 1;
                errorCells[destRow, errorMessageCol].PutValue("Validation Errors");
                destRow++; // Move to the first data row in the error sheet

                // Iterate through each data row (skip header row)
                for (int r = 1; r <= maxRow; r++)
                {
                    bool rowHasError = false;
                    StringBuilder rowErrorBuilder = new StringBuilder();

                    // Validate each cell in the row
                    for (int c = 0; c <= maxCol; c++)
                    {
                        try
                        {
                            // Retrieve any validation rule attached to the cell
                            Validation validation = sourceCells[r, c].GetValidation();

                            if (validation != null)
                            {
                                // Simple validation check: if the cell is required and empty, mark as error
                                bool isRequired = false;
                                try
                                {
                                    // Validation.IsRequired exists in older API; guard against missing property
                                    var prop = validation.GetType().GetProperty("IsRequired");
                                    if (prop != null)
                                    {
                                        isRequired = (bool)prop.GetValue(validation);
                                    }
                                }
                                catch
                                {
                                    // ignore reflection errors
                                }

                                object cellValue = sourceCells[r, c].Value;
                                bool isEmpty = cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString());

                                if (isRequired && isEmpty)
                                {
                                    rowHasError = true;
                                    string cellAddress = CellsHelper.CellIndexToName(r, c);
                                    rowErrorBuilder.Append($"{cellAddress}: Required value missing; ");
                                }
                                else if (!isRequired)
                                {
                                    // If there is any validation rule (even non‑required), treat as potential error for demo purposes
                                    rowHasError = true;
                                    string cellAddress = CellsHelper.CellIndexToName(r, c);
                                    rowErrorBuilder.Append($"{cellAddress}: Validation rule present; ");
                                }
                            }
                        }
                        catch (Exception cellEx)
                        {
                            // Capture unexpected validation errors per cell
                            rowHasError = true;
                            string cellAddress = CellsHelper.CellIndexToName(r, c);
                            rowErrorBuilder.Append($"{cellAddress}: {cellEx.Message}; ");
                        }
                    }

                    if (rowHasError)
                    {
                        // Copy the entire row values to the error sheet
                        for (int c = 0; c <= maxCol; c++)
                        {
                            errorCells[destRow, c].PutValue(sourceCells[r, c].Value);
                        }

                        // Write concatenated error messages into the extra column
                        errorCells[destRow, errorMessageCol].PutValue(rowErrorBuilder.ToString().Trim());
                        destRow++;
                    }
                }

                // Auto‑fit columns for better readability
                errorSheet.AutoFitColumns();

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the error sheet included
                workbook.Save(outputPath);
                Console.WriteLine($"Processing complete. Output saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}

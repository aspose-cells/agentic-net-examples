// Title: Export Invalid Data‑Validation Rows to an Error Sheet with Aspose.Cells for .NET
// Description: Demonstrates how to add a whole‑number validation (10‑20) to a column, iterate through the data rows, copy the header and any rows that break the rule to a new worksheet called ErrorReport, and save the workbook.
// Keywords: Aspose.Cells | C# | Excel data validation | export invalid rows | error report worksheet | CopyRows method | filter rows by validation | Aspose.Cells .NET example
// Common Searches: Aspose.Cells copy rows that fail validation | C# generate error report for Excel validation | How to export rows outside validation range using Aspose.Cells | Aspose.Cells filter invalid data rows | Create error sheet for data validation in .NET
// Developer Intent: Generate an Excel workbook that isolates rows not satisfying a specified data‑validation rule into a separate error‑report worksheet.
// Use Cases: Audit entries that fall outside an allowed numeric range by moving them to an error sheet. | Provide users with a concise list of invalid rows for correction or review. | Automate creation of validation‑based error reports in server‑side .NET applications.
// AI Prompts: Write C# code using Aspose.Cells that adds a whole‑number validation (10‑20) to column B, scans rows B2:B5, and copies rows violating the rule to a new worksheet named "ErrorReport". | Explain step‑by‑step how to create an error‑report sheet in Aspose.Cells by applying data validation, iterating over cells, and using CopyRows to transfer invalid rows. | Generate a complete .NET example that produces an Excel file with a data sheet and an error sheet containing only rows with invalid Quantity values.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsErrorReport
{
    // Demonstrates how to add a whole‑number validation (10‑20) to a column, iterate through the data rows, copy the header and any rows that break the rule to a new worksheet called ErrorReport, and save the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------------------------------------
                // 1. Create a new workbook and get the first sheet
                // -------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // -------------------------------------------------
                // 2. Populate sample data (some rows will violate validation)
                // -------------------------------------------------
                // Header
                dataSheet.Cells["A1"].PutValue("ID");
                dataSheet.Cells["B1"].PutValue("Quantity");

                // Valid rows
                dataSheet.Cells["A2"].PutValue(1);
                dataSheet.Cells["B2"].PutValue(15);   // within 10‑20

                dataSheet.Cells["A3"].PutValue(2);
                dataSheet.Cells["B3"].PutValue(8);    // below 10 → invalid

                dataSheet.Cells["A4"].PutValue(3);
                dataSheet.Cells["B4"].PutValue(25);   // above 20 → invalid

                dataSheet.Cells["A5"].PutValue(4);
                dataSheet.Cells["B5"].PutValue(12);   // valid

                // -------------------------------------------------
                // 3. Add a data‑validation rule: whole number between 10 and 20
                // -------------------------------------------------
                Validation validation = dataSheet.Validations[dataSheet.Validations.Add(new CellArea
                {
                    StartRow = 1,   // B2 (zero‑based)
                    StartColumn = 1,
                    EndRow = 4,     // B5
                    EndColumn = 1
                })];
                validation.Type = ValidationType.WholeNumber;
                validation.Operator = OperatorType.Between;
                validation.Formula1 = "10";
                validation.Formula2 = "20";
                validation.ShowError = true; // show error message when invalid data entered

                // -------------------------------------------------
                // 4. Create a new worksheet to hold error rows
                // -------------------------------------------------
                Worksheet errorSheet = workbook.Worksheets.Add("ErrorReport");

                // Copy header row to the error sheet (row 0)
                // Use the overload that copies between worksheets
                errorSheet.Cells.CopyRows(dataSheet.Cells, 0, 1, 0, null);

                // Keep track of the next row index in the error sheet (start after header)
                int errorRowIndex = 1;

                // -------------------------------------------------
                // 5. Manually check each data row against the validation rule
                // -------------------------------------------------
                for (int row = 1; row <= 4; row++) // rows B2:B5 (zero‑based)
                {
                    Cell qtyCell = dataSheet.Cells[row, 1]; // column B
                    if (qtyCell.Type != CellValueType.IsNumeric)
                        continue; // skip non‑numeric values

                    double qty = qtyCell.DoubleValue;
                    if (qty < 10 || qty > 20) // violates the rule
                    {
                        // Copy the entire offending row to the error sheet
                        errorSheet.Cells.CopyRows(dataSheet.Cells, row, 1, errorRowIndex, null);
                        errorRowIndex++;
                    }
                }

                // -------------------------------------------------
                // 6. Save the workbook
                // -------------------------------------------------
                string outputPath = "ErrorReportDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

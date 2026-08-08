// Title: C# – Preserve Quote Prefix and Add Whole‑Number Validation with StyleFlag in Aspose.Cells
// Description: Shows how to create a workbook, write values with a leading apostrophe, enable QuotePrefix via a StyleFlag, and apply a whole‑number (1‑100) data‑validation rule to the same range using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | QuotePrefix | StyleFlag | data validation | whole number validation | Excel leading apostrophe | .NET Excel library | range styling | CellArea Add validation
// Common Searches: Aspose.Cells preserve leading apostrophe | QuotePrefix style flag C# example | Add data validation to cells with QuotePrefix Aspose | How to apply QuotePrefix and validation in .NET | Aspose.Cells whole number validation sample
// Developer Intent: The developer wants to keep the leading quote (apostrophe) in cells while adding a numeric validation rule to the same range.
// Use Cases: Store product codes that start with an apostrophe but restrict user entry to numbers between 1 and 100. | Create a template where ID fields appear as text (with leading quote) yet must contain numeric values within a defined range. | Generate a report that displays textual identifiers while enforcing numeric input constraints on those cells.
// AI Prompts: Write C# code that applies a QuotePrefix style to a cell range and then adds a whole‑number validation rule using Aspose.Cells. | Explain how StyleFlag and QuotePrefix interact to retain leading apostrophes when applying data validation in Aspose.Cells for .NET. | Provide a step‑by‑step tutorial for creating a workbook, preserving leading quotes, and enforcing numeric input (1‑100) on the same cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixValidationDemo
{
    // Shows how to create a workbook, write values with a leading apostrophe, enable QuotePrefix via a StyleFlag, and apply a whole‑number (1‑100) data‑validation rule to the same range using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate a range with values that start with a single quote (quote prefix)
                // These values will be treated as text in Excel
                cells["A1"].PutValue("'001");
                cells["A2"].PutValue("'002");
                cells["A3"].PutValue("'003");

                // Create a style and enable QuotePrefix
                Style quoteStyle = workbook.CreateStyle();
                quoteStyle.QuotePrefix = true;

                // Create a StyleFlag and enable the QuotePrefix flag
                StyleFlag flag = new StyleFlag();
                flag.QuotePrefix = true;

                // Apply the style with the flag to the range A1:A3
                Aspose.Cells.Range range = cells.CreateRange("A1:A3");
                range.ApplyStyle(quoteStyle, flag);

                // Add a data validation rule to the same range (whole number between 1 and 100)
                // Use the newer Add(CellArea) overload
                CellArea area = CellArea.CreateCellArea(0, 0, 2, 0); // rows 0-2, column 0 (A)
                int validationIndex = sheet.Validations.Add(area);
                Validation validation = sheet.Validations[validationIndex];
                validation.Type = ValidationType.WholeNumber;
                validation.Operator = OperatorType.Between;
                validation.Formula1 = "1";
                validation.Formula2 = "100";
                validation.ShowError = true;
                validation.ErrorTitle = "Invalid Input";
                validation.ErrorMessage = "Please enter a number between 1 and 100.";

                // Save the workbook
                string outputPath = "QuotePrefixValidationDemo.xlsx";
                // Ensure the directory exists
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

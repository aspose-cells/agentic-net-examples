// Title: Preserve Quote Prefix and Add Whole‑Number Validation with StyleFlag in Aspose.Cells for .NET
// Description: Demonstrates how to keep a leading single‑quote display on cells, apply only the QuotePrefix attribute using a StyleFlag, and attach a whole‑number (1‑100) data‑validation rule via a CellArea overload. The workbook is saved as an Excel file.
// Keywords: Aspose.Cells | C# | QuotePrefix | StyleFlag | data validation | CellArea | whole number validation | Excel single quote | preserve leading apostrophe | apply style to range
// Common Searches: Aspose.Cells keep leading single quote | apply QuotePrefix with StyleFlag | add numeric validation to range Aspose.Cells | CellArea validation example C# | preserve text format while adding validation
// Developer Intent: Retain the visible leading single‑quote on cell values while applying a style and enforcing a whole‑number validation rule on the same range.
// Use Cases: Create templates where IDs start with a quote but users must enter numbers within a defined range. | Export data that includes quoted strings as text and still require input validation for numeric edits. | Generate workbooks that display quoted strings correctly and prevent out‑of‑range entries in the same cells.
// AI Prompts: Generate C# code using Aspose.Cells to apply a QuotePrefix style to A1:A5 with a StyleFlag and then add a whole‑number validation (1‑100) using CellArea. | Explain why StyleFlag is needed to isolate the QuotePrefix property when styling cells that also have data validation. | Provide a step‑by‑step tutorial for preserving leading single quotes while adding numeric validation in an Aspose.Cells workbook.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsQuotePrefixValidationDemo
{
    // Demonstrates how to keep a leading single‑quote display on cells, apply only the QuotePrefix attribute using a StyleFlag, and attach a whole‑number (1‑100) data‑validation rule via a CellArea overload. The workbook is saved as an Excel file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate cells A1:A5 with values that start with a single quote.
                // The single quote forces Excel to treat the content as text.
                for (int row = 0; row < 5; row++)
                {
                    sheet.Cells[row, 0].PutValue("'12345");
                }

                // Create a style that has QuotePrefix enabled.
                Style quoteStyle = workbook.CreateStyle();
                quoteStyle.QuotePrefix = true;

                // Create a StyleFlag that indicates only the QuotePrefix property should be applied.
                StyleFlag flag = new StyleFlag();
                flag.QuotePrefix = true;

                // Apply the style with the flag to the range A1:A5.
                AsposeRange dataRange = sheet.Cells.CreateRange("A1:A5");
                dataRange.ApplyStyle(quoteStyle, flag);

                // Add a data validation rule to the same range (whole numbers between 1 and 100).
                // Use the newer overload that accepts a CellArea.
                CellArea area = CellArea.CreateCellArea(0, 0, 4, 0); // rows 0-4, column 0
                int validationIndex = sheet.Validations.Add(area);
                Validation validation = sheet.Validations[validationIndex];
                validation.Type = ValidationType.WholeNumber;
                validation.Operator = OperatorType.Between;
                validation.Formula1 = "1";
                validation.Formula2 = "100";

                // Save the workbook.
                string outputPath = "QuotePrefixWithValidation.xlsx";
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

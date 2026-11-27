// Title: Preserving the leading apostrophe (QuotePrefix) while adding whole-number data validation to a cell range using StyleFlag in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# preserve leading apostrophe after adding data validation | apply QuotePrefix style flag to a range and set whole number validation in .NET | keep Excel cells with leading zeros as text when using Aspose.Cells validation | how to retain leading quote in cells A1:A5 while adding validation with Aspose.Cells | set QuotePrefix and add between 100 and 200 validation using Aspose.Cells C#
// Tags: QuotePrefix style flag Aspose.Cells | data validation whole number Aspose.Cells C# | apply style to range Aspose.Cells | preserve leading zeros Excel | Aspose.Cells range validation C#

using System;
using System.IO;
using Aspose.Cells;

namespace PreserveQuotePrefixWithValidation
{
    // The example creates a workbook, writes zero‑padded strings to cells A1‑A5, applies a style with QuotePrefix enabled via a StyleFlag so the cells display a leading apostrophe, adds a whole‑number validation (100–200) to the same range, saves the file, reloads it, and verifies that the QuotePrefix flag remains set.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // 1. Populate cells with values that should retain a leading quote
                // ------------------------------------------------------------
                // Cells A1 to A5 will contain text that looks like numbers.
                // We will set the QuotePrefix flag so Excel treats them as text.
                for (int i = 0; i < 5; i++)
                {
                    // Example value: 00123, 00124, ...
                    string value = (123 + i).ToString("D5"); // ensures leading zeros
                    Cell cell = worksheet.Cells[i, 0]; // column A
                    cell.PutValue(value); // put the raw string (no leading quote in the string)
                }

                // Create a style that enables QuotePrefix
                Style quoteStyle = workbook.CreateStyle();
                quoteStyle.QuotePrefix = true; // tells Excel to prepend a single quote when displaying

                // Apply the style only for the QuotePrefix flag to the range A1:A5
                Aspose.Cells.Range targetRange = worksheet.Cells.CreateRange("A1:A5");
                StyleFlag quoteFlag = new StyleFlag
                {
                    QuotePrefix = true // only the QuotePrefix property will be applied
                };
                targetRange.ApplyStyle(quoteStyle, quoteFlag);

                // ------------------------------------------------------------
                // 2. Add a data validation rule to the same range (A1:A5)
                // ------------------------------------------------------------
                // Define the area to which the validation applies (A1:A5)
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    EndRow = 4,
                    StartColumn = 0,
                    EndColumn = 0
                };

                // Add a validation for the defined area
                int validationIndex = worksheet.Validations.Add(area);
                Validation validation = worksheet.Validations[validationIndex];

                // The rule will allow only whole numbers between 100 and 200.
                validation.Type = ValidationType.WholeNumber;
                validation.Operator = OperatorType.Between;
                validation.Formula1 = "100";
                validation.Formula2 = "200";
                validation.InputMessage = "Enter a whole number between 100 and 200.";
                validation.ErrorMessage = "The value is not within the allowed range.";
                // Optional: show messages (default is true in most versions)
                // validation.ShowInputMessage = true;
                // validation.ShowErrorMessage = true;

                // ------------------------------------------------------------
                // 3. Save the workbook
                // ------------------------------------------------------------
                string outputPath = "QuotePrefix_With_Validation.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");

                // ------------------------------------------------------------
                // 4. Verify that QuotePrefix is still set after saving
                // ------------------------------------------------------------
                if (File.Exists(outputPath))
                {
                    Workbook loaded = new Workbook(outputPath);
                    Worksheet loadedSheet = loaded.Worksheets[0];

                    Console.WriteLine("\nVerification of QuotePrefix after reload:");
                    for (int i = 0; i < 5; i++)
                    {
                        Cell loadedCell = loadedSheet.Cells[i, 0];
                        bool isQuote = loadedCell.GetStyle().QuotePrefix;
                        Console.WriteLine($"Cell {loadedCell.Name}: QuotePrefix = {isQuote}, Value = {loadedCell.StringValue}");
                    }
                }
                else
                {
                    Console.WriteLine($"Error: The file '{outputPath}' was not found after saving.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            // Keep console window open when run from IDE
            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
        }
    }
}

// Title: Preserve leading single‑quote in Excel cells and add whole‑number validation using StyleFlag in Aspose.Cells for .NET
// AI Prompts: Write C# code that sets the QuotePrefix style on a range of cells with a StyleFlag and then creates a whole‑number (1‑999) data validation rule for the same range using Aspose.Cells. | Show how to apply a StyleFlag‑based QuotePrefix style to cells A1:A5 and attach a between‑operator numeric validation in an Aspose.Cells workbook.
// Common Searches: Aspose.Cells C# apply QuotePrefix style to a range and add numeric validation | How to keep leading apostrophe in Excel cells while using data validation with Aspose.Cells | Using StyleFlag to set QuotePrefix and validation for cells A1:A5 in .NET | Preserve text prefix and enforce whole number range in Aspose.Cells workbook | Apply QuotePrefix and whole number between validation in C# Aspose.Cells example
// Tags: QuotePrefix style flag Aspose.Cells | apply StyleFlag to cell range | data validation whole number Aspose.Cells | preserve leading single quote Excel .NET | range style and validation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsQuotePrefixValidationDemo
{
    // The example creates a workbook, fills cells A1:A5 with values that start with a single quote, uses a Style and StyleFlag to enable QuotePrefix for the range, adds a whole‑number (1‑999) data validation rule to the same cells, saves the file, reloads it, and confirms that the QuotePrefix flag remains set.
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

                // Fill cells A1:A5 with values that start with a single quote (treated as text)
                for (int i = 0; i < 5; i++)
                {
                    // Example value: '001, '002, etc.
                    cells[i, 0].PutValue($"'{(i + 1).ToString("D3")}");
                }

                // Create a style and enable QuotePrefix so the leading single quote is preserved in Excel
                Style quoteStyle = workbook.CreateStyle();
                quoteStyle.QuotePrefix = true; // Style property

                // Create a StyleFlag and enable the QuotePrefix flag
                StyleFlag flag = new StyleFlag();
                flag.QuotePrefix = true; // Flag property

                // Apply the style to the range A1:A5 using the flag
                AsposeRange range = cells.CreateRange(0, 0, 5, 1); // rows 0-4, column 0
                range.ApplyStyle(quoteStyle, flag); // Apply only the QuotePrefix setting

                // Define the area for validation (A1:A5)
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    EndRow = 4,
                    StartColumn = 0,
                    EndColumn = 0
                };

                // Add a data validation rule to the same range (whole number between 1 and 999)
                int validationIndex = sheet.Validations.Add(area);
                Validation validation = sheet.Validations[validationIndex];
                validation.Type = ValidationType.WholeNumber;
                validation.Operator = OperatorType.Between;
                validation.Formula1 = "1";
                validation.Formula2 = "999";

                // Save the workbook
                string filePath = "QuotePrefixValidationDemo.xlsx";
                workbook.Save(filePath);

                // Reload to verify that QuotePrefix is still set
                if (File.Exists(filePath))
                {
                    try
                    {
                        Workbook loaded = new Workbook(filePath);
                        Worksheet loadedSheet = loaded.Worksheets[0];
                        for (int i = 0; i < 5; i++)
                        {
                            bool qp = loadedSheet.Cells[i, 0].GetStyle().QuotePrefix;
                            Console.WriteLine($"Cell {i + 1} QuotePrefix: {qp}");
                        }
                    }
                    catch (Exception loadEx)
                    {
                        Console.WriteLine($"Error loading workbook: {loadEx.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Error: File '{filePath}' was not found after saving.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

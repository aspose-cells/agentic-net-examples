// Title: Localize numeric cells to German format using Aspose.Cells for .NET
// Description: Creates a workbook with mixed data, applies a generic number format, sets the workbook culture to German (de‑DE), detects numeric cells, converts the custom format by swapping commas and periods, assigns the culture‑specific pattern via Style.CultureCustom, and saves the file with localized number displays while leaving text cells untouched.
// Keywords: Aspose.Cells | C# number localization | German number format | CultureCustom | Cell.DisplayStringValue | NumberCategoryType | Excel locale formatting | replace decimal separator | thousand separator Aspose
// Common Searches: Aspose.Cells apply German number format | C# change decimal separator in Excel cells | detect numeric cells Aspose.Cells | localize Excel numbers to de-DE | use Style.CultureCustom Aspose
// Developer Intent: Apply a German‑style numeric format to every numeric cell in a worksheet while preserving non‑numeric content.
// Use Cases: Generate Excel reports that automatically display numbers using the target locale (e.g., German) without altering underlying values. | Convert existing custom number formats to match a specific culture for international distribution. | Create a reusable routine that localizes numeric displays across workbooks for multi‑regional applications.
// AI Prompts: Write C# code with Aspose.Cells that iterates over a used range, identifies numeric cells, and applies a de‑DE custom number format using Style.CultureCustom. | Explain how Cell.DisplayStringValue and NumberCategoryType can be leveraged to reformat numbers for a specific locale while keeping original data intact. | Provide a step‑by‑step method to transform a generic custom format (e.g., "#,##0.00") into a culture‑specific pattern by swapping commas and periods.

using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsNumberLocalization
{
    // Creates a workbook with mixed data, applies a generic number format, sets the workbook culture to German (de‑DE), detects numeric cells, converts the custom format by swapping commas and periods, assigns the culture‑specific pattern via Style.CultureCustom, and saves the file with localized number displays while leaving text cells untouched.
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

                // Sample data: mix of numeric and text values
                cells["A1"].PutValue(1234.56);          // numeric
                cells["A2"].PutValue("Sample Text");   // text
                cells["A3"].PutValue(987654321);       // numeric
                cells["A4"].PutValue(3.14159);         // numeric
                cells["A5"].PutValue("2021-12-31");    // date string (will stay as text)

                // Set a generic number format for all numeric cells (e.g., two decimal places)
                Style baseStyle = workbook.CreateStyle();
                baseStyle.Custom = "#,##0.00";
                StyleFlag flag = new StyleFlag { NumberFormat = true };
                Aspose.Cells.Range usedRange = cells.MaxDisplayRange;
                usedRange.ApplyStyle(baseStyle, flag);

                // Define the target culture for localization (German - uses '.' as thousand separator and ',' as decimal)
                CultureInfo targetCulture = new CultureInfo("de-DE");
                workbook.Settings.CultureInfo = targetCulture;

                // Iterate through each cell in the used range
                foreach (Cell cell in usedRange)
                {
                    // Retrieve the display string (formatted according to the cell's style)
                    string displayText = cell.DisplayStringValue;

                    // Detect if the cell currently holds a numeric value
                    if (cell.NumberCategoryType == NumberCategoryType.Number)
                    {
                        // Obtain the cell's current style
                        Style style = cell.GetStyle();

                        // Create a culture‑dependent custom format based on the target culture.
                        // For simplicity we reuse the existing pattern but replace separators.
                        string culturePattern = style.Custom;
                        if (!string.IsNullOrEmpty(culturePattern))
                        {
                            // Replace comma with temporary token, then replace dot, then restore token
                            culturePattern = culturePattern.Replace(",", "__TMP__")
                                                           .Replace(".", ",")
                                                           .Replace("__TMP__", ".");
                        }
                        else
                        {
                            // Fallback pattern if none is set
                            culturePattern = "#.##0,00";
                        }

                        // Assign the culture‑specific pattern
                        style.CultureCustom = culturePattern;

                        // Apply the updated style back to the cell (only number format changes)
                        cell.SetStyle(style, flag);

                        // Output the before/after values for verification
                        Console.WriteLine($"Cell {cell.Name}: Original='{displayText}' => Localized='{cell.DisplayStringValue}'");
                    }
                    else
                    {
                        // Non‑numeric cells are left unchanged; just display their content
                        Console.WriteLine($"Cell {cell.Name}: Text='{displayText}'");
                    }
                }

                // Ensure the output directory exists
                string outputPath = "LocalizedNumbers.xlsx";
                string outputDir = System.IO.Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !System.IO.Directory.Exists(outputDir))
                {
                    System.IO.Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to verify the applied formats
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

// Title: C# – Convert HTML to Excel and Apply Scientific Notation Formatting with Aspose.Cells
// Description: Load an HTML file using Aspose.Cells (LoadFormat.Html), convert any numeric strings—including scientific notation—to real numbers, apply the custom format "0.00E+00", and save the result as an XLSX workbook.
// Keywords: Aspose.Cells | HTML to Excel | C# scientific notation | custom number format | LoadFormat.Html | parse numeric strings | cell style | Excel export | numeric conversion | Aspose.Cells example
// Common Searches: Aspose.Cells load HTML keep numbers as numeric | C# apply scientific notation format after HTML import | convert HTML table to Excel with custom number format | parse scientific notation strings in Aspose.Cells | how to set custom number format in Aspose.Cells C#
// Developer Intent: Read an HTML document into a Workbook, turn numeric strings (including scientific notation) into true numeric cells, apply a custom scientific‑notation style, and export to XLSX.
// Use Cases: Convert web‑based financial reports that list large values in HTML to Excel, displaying all numbers in scientific notation. | Import engineering measurement tables from HTML where values appear as "1.23E+04" strings and need consistent numeric formatting. | Automate batch processing of multiple HTML files, standardizing numeric display across all worksheets with a custom format.
// AI Prompts: Generate C# code using Aspose.Cells to load an HTML file, detect cells containing numeric strings (including scientific notation), convert them to doubles, apply the custom format "0.00E+00", and save as XLSX. | Write a method that iterates through all used cells in an Aspose.Cells worksheet, parses string values to double with invariant culture, and sets a scientific‑notation style on numeric cells. | Provide an example of loading HTML with LoadOptions, handling both numeric and string cell types, applying a custom number format for scientific notation, and exporting the workbook to Excel.

using System;
using System.Globalization;
using Aspose.Cells;

// Load an HTML file using Aspose.Cells (LoadFormat.Html), convert any numeric strings—including scientific notation—to real numbers, apply the custom format "0.00E+00", and save the result as an XLSX workbook.
class HtmlToExcelWithScientificNotation
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlPath = "input.html";

        // Load the HTML file into a workbook
        // LoadOptions with LoadFormat.Html ensures proper parsing of HTML content
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
        Workbook workbook = new Workbook(htmlPath, loadOptions);

        // Get the first worksheet (or iterate through all worksheets if needed)
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Determine the used range of the worksheet
        int maxRow = cells.MaxDataRow;
        int maxColumn = cells.MaxDataColumn;

        // Custom number format for scientific notation (e.g., 1.23E+04)
        const string scientificFormat = "0.00E+00";

        // Iterate through each cell in the used range
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxColumn; col++)
            {
                Cell cell = cells[row, col];

                // If the cell already contains a numeric value, apply the custom format directly
                if (cell.Type == CellValueType.IsNumeric)
                {
                    Style style = cell.GetStyle();
                    style.Custom = scientificFormat;
                    cell.SetStyle(style);
                }
                // If the cell contains a string that can be parsed as a double (including scientific notation)
                else if (cell.Type == CellValueType.IsString)
                {
                    double numericValue;
                    // Try parsing using invariant culture to recognize scientific notation like "1.23E+04"
                    if (double.TryParse(cell.StringValue, NumberStyles.Float, CultureInfo.InvariantCulture, out numericValue))
                    {
                        // Replace the string with the numeric value
                        cell.PutValue(numericValue);

                        // Apply the scientific notation format
                        Style style = cell.GetStyle();
                        style.Custom = scientificFormat;
                        cell.SetStyle(style);
                    }
                }
            }
        }

        // Save the workbook as an Excel file (XLSX format)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}

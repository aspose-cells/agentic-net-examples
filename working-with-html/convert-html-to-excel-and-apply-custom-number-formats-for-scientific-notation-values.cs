// Title: C# – Convert HTML to Excel and apply scientific notation formatting with Aspose.Cells
// Description: Loads an HTML file into an Aspose.Cells Workbook, converts parsable strings to numeric values, iterates the used range, and assigns a custom scientific notation format (0.00E+00) to every numeric cell before saving as XLSX.
// Keywords: Aspose.Cells HTML to Excel | C# convert HTML workbook | custom number format scientific notation | Convert string to numeric Aspose.Cells | Load HTML workbook .NET | Apply custom format Excel C# | Aspose.Cells number formatting
// Common Searches: how to load html into aspocells workbook c# | set scientific notation format for cells aspocells | convert string numbers to numeric values aspocells | html table to excel with custom number format | aspocells custom number format example
// Developer Intent: Import an HTML document, turn numeric strings into true numbers, and display those numbers in scientific notation in the resulting Excel file.
// Use Cases: Generate Excel reports from HTML tables where large values must appear in scientific notation. | Automate batch conversion of HTML files to Excel while preserving numeric data types and applying a uniform format. | Build a data‑processing pipeline that normalizes HTML‑derived numbers and formats them for downstream analytics.
// AI Prompts: Write C# code using Aspose.Cells to load an HTML file, convert numeric strings to numbers, apply the custom format "0.00E+00" to all numeric cells, and save as XLSX. | Explain how to traverse the used range of a worksheet in Aspose.Cells and set a scientific notation style on cells of type IsNumeric. | Provide a step‑by‑step guide for batch‑processing a folder of HTML files, converting each to Excel with numeric conversion and scientific notation formatting using Aspose.Cells.

using System;
using Aspose.Cells;

namespace HtmlToExcelConversion
{
    // Loads an HTML file into an Aspose.Cells Workbook, converts parsable strings to numeric values, iterates the used range, and assigns a custom scientific notation format (0.00E+00) to every numeric cell before saving as XLSX.
    class Program
    {
        static void Main()
        {
            // Paths for source HTML and destination Excel files
            string htmlFile = "input.html";
            string excelFile = "output.xlsx";

            // Load the HTML file into a workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(htmlFile, loadOptions);

            // Get the first worksheet's cells collection
            Cells cells = workbook.Worksheets[0].Cells;

            // Convert any string that can be interpreted as a number to a numeric value
            cells.ConvertStringToNumericValue();

            // Apply a custom scientific notation format to all numeric cells
            // Define the format pattern (e.g., two decimal places in scientific notation)
            string scientificFormat = "0.00E+00";

            // Iterate through the used range of cells
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        // Retrieve the existing style, modify the custom format, and reapply
                        Style style = cell.GetStyle();
                        style.Custom = scientificFormat;
                        cell.SetStyle(style);
                    }
                }
            }

            // Save the workbook as an Excel file
            workbook.Save(excelFile, SaveFormat.Xlsx);

            Console.WriteLine("HTML has been converted to Excel with scientific notation formatting.");
        }
    }
}

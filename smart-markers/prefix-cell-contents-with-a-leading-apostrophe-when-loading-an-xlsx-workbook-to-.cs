using System;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Load the workbook (existing XLSX)
            Workbook workbook = new Workbook(sourcePath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the cells collection of the current worksheet
                Cells cells = sheet.Cells;

                // Loop through all used cells
                foreach (Cell cell in cells)
                {
                    // Process only string cells (skip numbers, dates, etc.)
                    if (cell.Type == CellValueType.IsString)
                    {
                        // Retrieve the current style
                        Style style = cell.GetStyle();

                        // Enable QuotePrefix to indicate the value starts with a leading apostrophe
                        style.QuotePrefix = true;

                        // Apply the modified style back to the cell
                        cell.SetStyle(style);
                    }
                }
            }

            // Save the modified workbook; the cells will now display with a leading apostrophe in Excel
            string outputPath = "output_with_quoteprefix.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to '{outputPath}' with QuotePrefix applied to string cells.");
        }
    }
}
// Title: Insert a Year‑Over‑Year Growth % Column into an Excel Sheet Using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that adds a new column named "YoY Growth %" and fills each cell with a formula that computes the percentage change between the current year's sales and the previous year's sales for the same product. | Generate an IFERROR‑wrapped SUMIFS expression in Aspose.Cells to calculate year‑over‑year growth and assign it to the newly created column for every data row.
// Common Searches: asp.net add YoY growth column to existing Excel file using Aspose.Cells | Aspose.Cells C# formula for year over year sales growth per product | how to use SUMIFS in Aspose.Cells to reference previous year values | set cell formula dynamically for each row with Aspose.Cells C#
// Tags: insert calculated column Aspose.Cells C# | YoY growth percentage formula Aspose.Cells | SUMIFS IFERROR expression Aspose.Cells | set cell formula programmatically .NET Excel | calculate year over year sales growth C#

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, inserts a "YoY Growth %" header in a new column, and for each data row assigns an IFERROR‑wrapped SUMIFS formula that calculates the year‑over‑year sales growth for the product, then saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define column indexes (zero‑based)
            int productCol = 0; // Column A
            int yearCol = 1;    // Column B
            int salesCol = 2;   // Column C
            int growthCol = 3;  // Column D (new column)

            // Add header for the YoY Growth % column
            cells[0, growthCol].PutValue("YoY Growth %");

            // Determine the last row with data (zero‑based)
            int lastRow = cells.MaxDataRow;

            // Loop through each data row and set the YoY growth formula
            for (int row = 1; row <= lastRow; row++)
            {
                // Build cell addresses (e.g., A2, B2, C2) using the Cells collection
                string productCell = cells[row, productCol].Name; // e.g., A2
                string yearCell = cells[row, yearCol].Name;       // e.g., B2
                string salesCell = cells[row, salesCol].Name;     // e.g., C2

                // Formula:
                // IFERROR((CurrentSales - SUMIFS(C:C, A:A, Product, B:B, Year-1)) /
                //         SUMIFS(C:C, A:A, Product, B:B, Year-1), 0)
                string formula = $"IFERROR(({salesCell}-SUMIFS(C:C, A:A, {productCell}, B:B, {yearCell}-1))/SUMIFS(C:C, A:A, {productCell}, B:B, {yearCell}-1),0)";

                // Assign the formula to the YoY Growth % cell
                cells[row, growthCol].Formula = formula;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

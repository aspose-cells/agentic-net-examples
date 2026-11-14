// Title: How to apply a custom currency number format to a column in an Aspose.Cells ListObject (Excel table) using C#
// Common Searches: C# Aspose.Cells set currency format for a table column | Apply custom number format to ListObject column in Excel with Aspose.Cells | Formatting price column as $#,##0.00 in Aspose.Cells workbook | How to use StyleFlag to format cells in an Aspose.Cells table | Aspose.Cells C# example for currency formatting in a ListObject
// Tags: Aspose.Cells currency style C# | ListObject column styling Aspose.Cells | StyleFlag NumberFormat Aspose.Cells | Excel table column formatting C# | Set cell style programmatically Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a workbook, defines a range as a ListObject (Excel table) with "Item" and "Price" columns, then builds a custom currency style ("$#,##0.00") and applies it to every cell in the "Price" column using a Style and StyleFlag before saving the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add headers
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Price");

            // Add sample numeric data
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.5);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.75);
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(2.25);

            // Define the range that will become a table (including header row)
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 5;         // rows 1‑5 (header + 4 data rows)
            int totalColumns = 2;      // columns A‑B

            // Create a ListObject (table) over the defined range
            int tableIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "Products";

            // Apply a custom currency number format to the "Price" column
            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Custom = "$#,##0.00";   // custom format for currency
            StyleFlag flag = new StyleFlag { NumberFormat = true };

            // Worksheet column index for the "Price" column (B = index 1)
            int priceColumnIndex = firstColumn + 1;

            // Apply the style to each data cell in the "Price" column
            int firstDataRow = firstRow + 1; // first row after header
            int lastDataRow = firstRow + totalRows - 1;
            for (int row = firstDataRow; row <= lastDataRow; row++)
            {
                Cell cell = sheet.Cells[row, priceColumnIndex];
                cell.SetStyle(currencyStyle, flag);
            }

            // Ensure output directory exists
            string outputPath = "Output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

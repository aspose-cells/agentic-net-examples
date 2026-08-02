// Title: Apply Currency Number Format to a Table Column with Aspose.Cells for .NET
// Description: Creates a workbook, adds a ListObject (table) over A1:B4, defines a "$#,##0.00" style, selects the Price column cells, and applies only the number format using StyleFlag before saving the file.
// Keywords: Aspose.Cells | C# | .NET | custom number format | currency format | table column styling | ListObject | StyleFlag | Excel export | financial reporting
// Common Searches: Aspose.Cells set currency format for table column | C# apply custom number format to ListObject column | StyleFlag only number format Aspose.Cells | format price column as $ in Excel using Aspose
// Developer Intent: Display numeric values in a table column as formatted currency without altering other cell styles.
// Use Cases: Generate product price lists where the Price column shows dollar values with two decimals. | Create financial statements that automatically format monetary columns for readability. | Export invoice data to Excel with the amount column pre‑formatted as currency.
// AI Prompts: Show C# code that applies the "$#,##0.00" number format to a specific column of an Aspose.Cells ListObject using StyleFlag. | Explain how to format only the numeric cells of a table column as currency while preserving existing cell formatting in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a ListObject (table) over A1:B4, defines a "$#,##0.00" style, selects the Price column cells, and applies only the number format using StyleFlag before saving the file.
    public class TableColumnCurrencyFormatDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the table
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.25);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.75);
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(2.50);

            // Create a table (ListObject) that includes the data range A1:B4
            int tableIndex = sheet.ListObjects.Add(0, 0, 3, 1, true);
            ListObject table = sheet.ListObjects[tableIndex];
            // Set table name (use DisplayName if Name property is unavailable)
            table.DisplayName = "ProductsTable";

            // Define a custom currency number format
            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Custom = "$#,##0.00";

            // Determine the data rows range for the Price column (excluding header)
            int firstDataRow = table.DataRange.FirstRow + 1; // row after header
            int lastDataRow = table.DataRange.FirstRow + table.DataRange.RowCount - 1;
            int columnIndex = table.DataRange.FirstColumn + 1; // Price column (B)

            // Create the range for the price column data cells
            Aspose.Cells.Range priceRange = sheet.Cells.CreateRange(
                firstDataRow,
                columnIndex,
                lastDataRow - firstDataRow + 1,
                1);

            // Apply only the number format using StyleFlag
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;
            priceRange.ApplyStyle(currencyStyle, flag);

            // Save the workbook
            string outputPath = "TableColumnCurrencyFormatDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}

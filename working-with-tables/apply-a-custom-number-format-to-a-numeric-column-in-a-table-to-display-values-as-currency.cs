// Title: Format a Table Column as Currency with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a ListObject table, defines a style with the custom format "$#,##0.00", uses a StyleFlag to apply only the number‑format to the price column, and saves the file as an XLSX workbook.
// Keywords: Aspose.Cells | C# | custom number format | currency format | ListObject | ListColumn | ApplyStyle | StyleFlag | Excel table column formatting | Aspose.Cells .NET
// Common Searches: Aspose.Cells set currency format for table column | C# apply custom number format to ListColumn | StyleFlag only number format Aspose.Cells | How to format price column in Excel table using Aspose.Cells | ApplyStyle currency format ListObject column
// Developer Intent: Apply a custom currency number format to the numeric column of an Excel table without affecting other cell styles.
// Use Cases: Generate product price lists where the price column shows values with a dollar sign and two decimals. | Produce financial reports that automatically display amount columns as currency. | Export invoices from an application with the total column pre‑formatted as currency in the resulting Excel file.
// AI Prompts: Show how to format a ListColumn as currency in Aspose.Cells while preserving other styles. | Provide a C# example that uses StyleFlag to change only the number format of a table column. | Explain the steps to create a custom currency style and apply it to a specific column of a ListObject in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a ListObject table, defines a style with the custom format "$#,##0.00", uses a StyleFlag to apply only the number‑format to the price column, and saves the file as an XLSX workbook.
    public class TableColumnCurrencyFormat
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data: first column is product name, second column is price
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Price");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1.25);
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["B3"].PutValue(0.85);
                sheet.Cells["A4"].PutValue("Banana");
                sheet.Cells["B4"].PutValue(0.60);

                // Create a table that includes the data range (including headers)
                int tableIndex = sheet.ListObjects.Add(0, 0, 4, 1, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "ProductsTable";

                // Prepare a style with a custom currency number format
                Style currencyStyle = workbook.CreateStyle();
                currencyStyle.Custom = "$#,##0.00";

                // Use StyleFlag to apply only the number format part of the style
                StyleFlag flag = new StyleFlag();
                flag.NumberFormat = true;

                // Apply the style to the numeric column (second column, index 1) of the table
                ListColumn priceColumn = table.ListColumns[1];
                priceColumn.Range.ApplyStyle(currencyStyle, flag);

                // Save the workbook to a file
                workbook.Save("TableColumnCurrencyFormat.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            TableColumnCurrencyFormat.Run();
        }
    }
}

// Title: Apply a Custom Accounting Currency Format to Smart‑Marker Populated Cells in Aspose.Cells for .NET
// Description: This example creates a workbook, inserts smart markers for product names and amounts, processes a List&lt;ProductInfo&gt; to expand the markers, then defines and applies a custom accounting currency style ("_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)") only to the number‑format of the populated amount column before saving the file.
// Keywords: Aspose.Cells | C# | .NET | smart markers | custom number format | currency formatting | accounting style | Excel export | WorkbookDesigner | StyleFlag
// Common Searches: Aspose.Cells set custom currency format after smart markers | C# apply accounting number format to smart‑marker generated cells | how to format numbers with parentheses for negatives in Aspose.Cells | custom number format for smart marker data in .NET
// Developer Intent: Add a custom accounting currency number format to the column that receives values from smart markers after the data is processed.
// Use Cases: Generate invoices where product prices are filled via smart markers and displayed with a consistent accounting currency format. | Create financial reports that expand transaction data with smart markers and automatically enforce a locale‑specific currency style on amount columns. | Export sales dashboards to Excel, populating dynamic data through smart markers and applying a uniform currency format to totals and subtotals.
// AI Prompts: Show how to change the custom format to use the Euro (€) symbol while keeping the accounting style. | Demonstrate applying the same custom currency format to multiple columns (e.g., Amount, Tax, Discount) after smart marker processing. | Explain how to determine the populated range dynamically when the data source size is unknown before processing.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerNumberFormat
{
    // Simple data class for demonstration
    // This example creates a workbook, inserts smart markers for product names and amounts, processes a List&lt;ProductInfo&gt; to expand the markers, then defines and applies a custom accounting currency style ("_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)") only to the number‑format of the populated amount column before saving the file.
    public class ProductInfo
    {
        public string Product { get; set; }
        public double Amount { get; set; }
    }

    public class ApplyCurrencyFormatWithSmartMarkers
    {
        public static void Run()
        {
            // -------------------------------------------------
            // 1. Create a workbook and set up smart markers
            // -------------------------------------------------
            Workbook workbook = new Workbook();                     // create workbook
            Worksheet sheet = workbook.Worksheets[0];               // first worksheet

            // Header row
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Amount");

            // Smart markers – they will be expanded when the designer processes the data source
            sheet.Cells["A2"].PutValue("&=Data.Product");
            sheet.Cells["B2"].PutValue("&=Data.Amount");

            // -------------------------------------------------
            // 2. Prepare data source and process smart markers
            // -------------------------------------------------
            List<ProductInfo> data = new List<ProductInfo>
            {
                new ProductInfo { Product = "Apple",  Amount = 1234.56 },
                new ProductInfo { Product = "Banana", Amount = 7890.12 },
                new ProductInfo { Product = "Cherry", Amount = 345.67 }
            };

            WorkbookDesigner designer = new WorkbookDesigner();     // create designer
            designer.Workbook = workbook;                           // assign workbook
            designer.SetDataSource("Data", data);                   // set data source
            designer.Process();                                     // populate smart markers

            // -------------------------------------------------
            // 3. Apply a custom currency number format to the populated amount column
            // -------------------------------------------------
            // Determine the range that now contains the amount values.
            // After processing, rows 2..(data.Count+1) in column B hold the numbers.
            int firstDataRow = 1; // zero‑based index (row 2 in Excel)
            int lastDataRow = firstDataRow + data.Count - 1;
            Aspose.Cells.Range amountRange = sheet.Cells.CreateRange(firstDataRow, 1, data.Count, 1);

            // Create a style with the desired custom currency format.
            Style currencyStyle = workbook.CreateStyle();
            // Accounting style with two decimal places, negative numbers in parentheses.
            currencyStyle.Custom = "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)";

            // Use StyleFlag to apply only the number format part of the style.
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;   // apply only number format

            // Apply the style to the range.
            amountRange.ApplyStyle(currencyStyle, flag);

            // -------------------------------------------------
            // 4. Save the result
            // -------------------------------------------------
            workbook.Save("SmartMarkerCurrencyFormatted.xlsx");
        }
    }

    // Entry point for testing
    class Program
    {
        static void Main()
        {
            ApplyCurrencyFormatWithSmartMarkers.Run();
            Console.WriteLine("Workbook created with currency formatting applied to smart‑marker populated cells.");
        }
    }
}

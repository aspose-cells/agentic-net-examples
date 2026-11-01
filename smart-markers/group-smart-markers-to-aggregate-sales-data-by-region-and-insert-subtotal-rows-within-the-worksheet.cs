// Title: How to group smart markers by region and insert subtotal rows in an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# group smart markers by region and add subtotal rows | Insert subtotal rows after processing smart markers with WorkbookDesigner in .NET | Generate sales report with region subtotals using Aspose.Cells smart markers | Calculate and style region subtotals in Excel using Aspose.Cells C# | Auto‑fit columns after adding subtotal rows in an Aspose.Cells workbook
// Tags: smart markers region grouping Aspose.Cells | WorkbookDesigner subtotal row insertion C# | styled subtotal row Excel Aspose.Cells | auto‑fit columns after subtotal Aspose.Cells | DataSet processing with smart markers Aspose.Cells

using System;
using System.Data;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;

// The example demonstrates creating a workbook template with smart markers, populating a DataTable of sales data, processing the template via WorkbookDesigner, then iterating the generated rows to calculate totals per region. It inserts a formatted subtotal row after each region, applies bold and gray styling, auto‑fits columns, and saves the result as GroupedSalesReport.xlsx.
class GroupSmartMarkersExample
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook (template) and add a worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SalesReport";

            // 2. Build the smart marker template.
            // Header row
            sheet.Cells["A1"].PutValue("Region");
            sheet.Cells["B1"].PutValue("Sales");

            // Collection marker (must be placed above the data rows)
            sheet.Cells["A2"].PutValue("&=SalesData");

            // Data row with field markers
            sheet.Cells["A3"].PutValue("&=Region");
            sheet.Cells["B3"].PutValue("&=Amount");

            // 3. Prepare the data source: a DataTable with Region and Amount columns.
            DataTable dt = new DataTable("SalesData");
            dt.Columns.Add("Region", typeof(string));
            dt.Columns.Add("Amount", typeof(double));

            // Sample data
            dt.Rows.Add("North", 1200.50);
            dt.Rows.Add("North", 850.75);
            dt.Rows.Add("South", 430.00);
            dt.Rows.Add("South", 780.25);
            dt.Rows.Add("South", 210.40);
            dt.Rows.Add("East", 950.00);
            dt.Rows.Add("East", 640.30);
            dt.Rows.Add("West", 300.00);
            dt.Rows.Add("West", 470.20);
            dt.Rows.Add("West", 125.80);

            // 4. Process the smart markers using WorkbookDesigner.
            DataSet ds = new DataSet();
            ds.Tables.Add(dt); // The table name matches the collection marker "SalesData"

            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(ds);
            designer.Process();

            // 5. After processing, insert subtotal rows for each region.
            int startRow = 2; // zero‑based index for row 3 (first data row)
            int currentRow = startRow;
            string currentRegion = null;
            double regionSum = 0.0;

            while (currentRow <= sheet.Cells.MaxDataRow)
            {
                string region = sheet.Cells[currentRow, 0].StringValue; // Column A (Region)
                double amount = sheet.Cells[currentRow, 1].DoubleValue; // Column B (Sales)

                if (currentRegion == null)
                {
                    currentRegion = region;
                    regionSum = amount;
                }
                else if (region == currentRegion)
                {
                    regionSum += amount;
                }
                else
                {
                    InsertSubtotalRow(sheet, currentRow, currentRegion, regionSum);
                    currentRow++; // skip the inserted subtotal row
                    currentRegion = region;
                    regionSum = amount;
                }

                currentRow++;
            }

            // Insert subtotal for the last region if any data was processed.
            if (currentRegion != null)
            {
                InsertSubtotalRow(sheet, currentRow, currentRegion, regionSum);
            }

            // 6. Auto‑fit columns for better appearance.
            sheet.AutoFitColumns();

            // 7. Save the workbook.
            workbook.Save("GroupedSalesReport.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }

    /// <param name="sheet">Worksheet where the row will be inserted.</param>
    /// <param name="rowIndex">Zero‑based index where the subtotal row should appear.</param>
    /// <param name="region">Region name for the subtotal.</param>
    /// <param name="subtotal">Aggregated sales amount for the region.</param>
    private static void InsertSubtotalRow(Worksheet sheet, int rowIndex, string region, double subtotal)
    {
        // Insert a new row.
        sheet.Cells.InsertRows(rowIndex, 1);

        // Write the subtotal label and value.
        sheet.Cells[rowIndex, 0].PutValue(region + " Subtotal");
        sheet.Cells[rowIndex, 1].PutValue(subtotal);

        // Apply a simple style to highlight the subtotal row.
        Style style = sheet.Cells[rowIndex, 0].GetStyle();
        style.Font.IsBold = true;
        style.ForegroundColor = Color.LightGray;
        style.Pattern = BackgroundType.Solid;
        sheet.Cells.CreateRange(rowIndex, 0, 1, 2).ApplyStyle(style, new StyleFlag { All = true });
    }
}

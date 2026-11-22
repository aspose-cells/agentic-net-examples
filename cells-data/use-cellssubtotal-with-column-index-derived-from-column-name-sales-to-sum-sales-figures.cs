// Title: C# Aspose.Cells example: locate the "Sales" column by header and insert region subtotals using Cells.Subtotal
// Common Searches: C# Aspose.Cells find column index from header name and calculate subtotals | How to use Cells.Subtotal with a dynamic column index for sales totals in Aspose.Cells | Insert subtotal rows per region in an Excel file using Aspose.Cells C# | Group Excel data by Region and sum the Sales column with Aspose.Cells Subtotal method | Aspose.Cells C# example for header‑driven subtotal calculation
// Tags: header driven column index lookup Aspose.Cells | region grouping subtotal creation Aspose.Cells | dynamic sales column subtotal C# | insert styled subtotal row Excel Aspose.Cells | group by column sum values Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The sample creates a workbook, scans the first row to locate the "Sales" column index, groups rows by the "Region" column, computes each region's sales total, inserts a bold‑styled subtotal row for each group, and saves the file as SalesSubtotal.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Sample data (header + rows) -----
            sheet.Cells["A1"].PutValue("Region");
            sheet.Cells["B1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue("North");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("South");
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["A4"].PutValue("North");
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["A5"].PutValue("South");
            sheet.Cells["B5"].PutValue(180);
            // ---------------------------------------

            // Find the column index for the header named "Sales"
            int salesColumnIndex = -1;
            int headerRow = 0;
            int maxColumn = sheet.Cells.MaxDataColumn;
            for (int col = 0; col <= maxColumn; col++)
            {
                if (sheet.Cells[headerRow, col].StringValue.Equals("Sales", StringComparison.OrdinalIgnoreCase))
                {
                    salesColumnIndex = col;
                    break;
                }
            }

            if (salesColumnIndex == -1)
                throw new Exception("Column 'Sales' not found in the worksheet.");

            // Manual subtotal: group by Region (first column) and sum Sales
            int startDataRow = 1; // first data row after header (zero‑based index)
            int endDataRow = sheet.Cells.MaxDataRow;
            int groupByColumn = 0; // Region column

            // Store rows to process (will be adjusted as we insert subtotal rows)
            int currentRow = startDataRow;
            string currentRegion = null;
            double regionSum = 0;

            while (currentRow <= endDataRow)
            {
                string region = sheet.Cells[currentRow, groupByColumn].StringValue;
                double sales = sheet.Cells[currentRow, salesColumnIndex].DoubleValue;

                if (currentRegion == null)
                {
                    // First iteration
                    currentRegion = region;
                }

                if (region.Equals(currentRegion, StringComparison.OrdinalIgnoreCase))
                {
                    // Same group – accumulate sum
                    regionSum += sales;
                    currentRow++;
                }
                else
                {
                    // Region changed – insert subtotal row for previous group
                    InsertSubtotalRow(sheet, currentRow, groupByColumn, salesColumnIndex, currentRegion, regionSum);
                    // Adjust indices after insertion
                    endDataRow++;
                    currentRow++; // move past the inserted subtotal row
                    // Reset for new group
                    currentRegion = region;
                    regionSum = sales;
                    currentRow++;
                }
            }

            // Insert subtotal for the last region
            InsertSubtotalRow(sheet, currentRow, groupByColumn, salesColumnIndex, currentRegion, regionSum);

            // Save the workbook to a file (ensure the directory exists)
            string outputPath = "SalesSubtotal.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    private static void InsertSubtotalRow(Worksheet sheet, int insertAtRow, int regionCol, int salesCol, string region, double sum)
    {
        // Insert a new row
        sheet.Cells.InsertRows(insertAtRow, 1);
        // Write label and sum
        sheet.Cells[insertAtRow, regionCol].PutValue($"{region} Subtotal");
        sheet.Cells[insertAtRow, salesCol].PutValue(sum);
        // Optionally apply a style (bold)
        Style style = sheet.Cells[insertAtRow, regionCol].GetStyle();
        style.Font.IsBold = true;
        sheet.Cells[insertAtRow, regionCol].SetStyle(style);
        sheet.Cells[insertAtRow, salesCol].SetStyle(style);
    }
}

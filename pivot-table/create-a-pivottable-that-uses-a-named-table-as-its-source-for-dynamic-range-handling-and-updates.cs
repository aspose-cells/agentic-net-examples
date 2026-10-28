// Title: Create a PivotTable from a named ListObject (dynamic range) using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create pivot table from ListObject named table | how to use a named table as pivot source in Aspose.Cells .NET | dynamic range pivot table with Aspose.Cells ListObject example | refresh pivot table after adding data to Aspose.Cells ListObject | set currency format for data field in Aspose.Cells pivot table C#
// Tags: Aspose.Cells ListObject pivot source | dynamic range pivot table C# | pivot table currency formatting Aspose.Cells | refresh calculate pivot data Aspose.Cells | generate pivot table from ListObject Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;   // For ListObject and TableStyleType
using Aspose.Cells.Pivot;    // For PivotTable and PivotFieldType

// The example builds a new workbook, fills a worksheet with sample sales data, converts the range into a named ListObject called "SalesData", constructs a source address from the table, adds a PivotTable on a separate sheet using that address, configures Region as rows, Product as columns, and Sales as a summed data field with currency formatting, refreshes and calculates the pivot, and saves the file as "PivotTableWithNamedTable.xlsx".
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Prepare source data on the first worksheet
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Sample data (including header row)
            string[,] sampleData = new string[,]
            {
                { "Region", "Product", "Sales" },
                { "North",  "Apple",   "1200" },
                { "South",  "Banana",  "850"  },
                { "North",  "Banana",  "950"  },
                { "East",   "Apple",   "1300" },
                { "West",   "Orange",  "700"  },
                { "South",  "Apple",   "1150" },
                { "East",   "Banana",  "900"  },
                { "West",   "Apple",   "1050" }
            };

            // Fill the worksheet with the sample data
            for (int i = 0; i < sampleData.GetLength(0); i++)
            {
                for (int j = 0; j < sampleData.GetLength(1); j++)
                {
                    dataSheet.Cells[i, j].PutValue(sampleData[i, j]);
                }
            }

            // -------------------------------------------------
            // 2. Convert the range into a named table (ListObject)
            // -------------------------------------------------
            int totalRows = sampleData.GetLength(0);
            int totalCols = sampleData.GetLength(1);

            // Add a ListObject that covers the data range (including headers)
            int tableIndex = dataSheet.ListObjects.Add(0, 0, totalRows, totalCols, true);
            ListObject table = dataSheet.ListObjects[tableIndex];
            table.DisplayName = "SalesData";                     // Named table for dynamic range handling
            table.ShowHeaderRow = true;
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // -------------------------------------------------
            // 3. Create a worksheet that will host the PivotTable
            // -------------------------------------------------
            Worksheet pivotSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            pivotSheet.Name = "Pivot";

            // -------------------------------------------------
            // 4. Build the source address string from the named table
            // -------------------------------------------------
            // Range does not expose StartRow/StartColumn directly; use FirstRow/FirstColumn.
            int startRow = table.DataRange.FirstRow;
            int startCol = table.DataRange.FirstColumn;
            int endRow = startRow + table.DataRange.RowCount - 1;
            int endCol = startCol + table.DataRange.ColumnCount - 1;

            string startCell = CellsHelper.CellIndexToName(startRow, startCol);
            string endCell   = CellsHelper.CellIndexToName(endRow, endCol);
            string sourceAddress = $"{dataSheet.Name}!{startCell}:{endCell}";

            // -------------------------------------------------
            // 5. Add the PivotTable using the named table as source
            // -------------------------------------------------
            int pivotRow = 0;      // top‑left row of the PivotTable
            int pivotColumn = 0;   // top‑left column of the PivotTable
            int pivotIndex = pivotSheet.PivotTables.Add("PivotTable1", pivotRow, pivotColumn, sourceAddress);
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure fields:
            //   Row field  : Region
            //   Column field: Product
            //   Data field : Sales (Sum)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Optional: format the data field as currency
            if (pivotTable.DataFields.Count > 0)
            {
                pivotTable.DataFields[0].NumberFormat = "$#,##0";
            }

            // Refresh and calculate the PivotTable data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // -------------------------------------------------
            // 6. Save the workbook (lifecycle: save)
            // -------------------------------------------------
            string outputPath = "PivotTableWithNamedTable.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

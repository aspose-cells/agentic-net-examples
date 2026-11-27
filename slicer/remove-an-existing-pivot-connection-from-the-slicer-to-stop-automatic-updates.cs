// Title: Detach a slicer from its pivot table in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# remove slicer pivot connection example | How to disconnect a slicer from a pivot table with Aspose.Cells | C# code to stop slicer updates on a pivot table in Aspose.Cells | RemovePivotConnection method usage in Aspose.Cells slicer | Detach slicer from pivot table in Excel using Aspose.Cells .NET
// Tags: aspnet cells slicer remove pivot connection | c# aspose.cells detach slicer from pivot | aspose.cells removepivotconnection method | excel slicer pivot unlink aspose.cells | aspose.cells slicer manipulation c#

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace RemoveSlicerPivotConnectionDemo
{
    // The example creates a workbook, adds sample data, builds a pivot table, inserts a slicer linked to the pivot, then calls RemovePivotConnection to detach the slicer, preventing automatic pivot updates, and finally saves the Excel file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and populate it with sample data.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Header row
                dataSheet.Cells["A1"].PutValue("Product");
                dataSheet.Cells["B1"].PutValue("Region");
                dataSheet.Cells["C1"].PutValue("Sales");

                // Sample data
                string[,] sample = new string[,]
                {
                    { "Apple",  "North", "1200" },
                    { "Apple",  "South", "850"  },
                    { "Banana", "North", "950"  },
                    { "Banana", "South", "770"  },
                    { "Cherry", "North", "430"  },
                    { "Cherry", "South", "610"  }
                };

                for (int i = 0; i < sample.GetLength(0); i++)
                {
                    dataSheet.Cells[i + 2, 0].PutValue(sample[i, 0]); // Product
                    dataSheet.Cells[i + 2, 1].PutValue(sample[i, 1]); // Region
                    dataSheet.Cells[i + 2, 2].PutValue(Convert.ToDouble(sample[i, 2])); // Sales
                }

                // -----------------------------------------------------------------
                // 2. Add a PivotTable based on the data range.
                // -----------------------------------------------------------------
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
                PivotTableCollection pivotTables = pivotSheet.PivotTables;

                // Define the source range (including headers)
                string sourceRange = "=Data!A1:C7";

                // Destination cell for the pivot table
                string destCell = "A3";

                int pivotIndex = pivotTables.Add(sourceRange, destCell, "SalesPivot");
                PivotTable pivotTable = pivotTables[pivotIndex];

                // Configure the pivot fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh and calculate the pivot data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // -----------------------------------------------------------------
                // 3. Add a slicer that is initially connected to the pivot table.
                // -----------------------------------------------------------------
                // The correct overload is Add(PivotTable pivot, string destCellName, int baseFieldIndex)
                // "Product" is the first field (index 0) in the source data.
                int slicerIndex = pivotSheet.Slicers.Add(pivotTable, "E3", 0);
                Slicer slicer = pivotSheet.Slicers[slicerIndex];

                Console.WriteLine("Slicer created and connected to pivot table.");

                // -----------------------------------------------------------------
                // 4. Remove the pivot connection from the slicer.
                //    After this call the slicer will no longer trigger automatic
                //    updates of the pivot table when its selection changes.
                // -----------------------------------------------------------------
                slicer.RemovePivotConnection(pivotTable);
                Console.WriteLine("Pivot connection removed from slicer.");

                // -----------------------------------------------------------------
                // 5. Save the workbook to verify the result.
                // -----------------------------------------------------------------
                string outputPath = "RemovePivotConnectionDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
                Environment.ExitCode = 1;
            }
        }
    }
}

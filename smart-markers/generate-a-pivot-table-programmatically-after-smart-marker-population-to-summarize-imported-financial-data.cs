// Title: Create an Excel pivot table from smart‑marker populated financial records with Aspose.Cells in C#
// AI Prompts: Write C# code that defines smart markers, binds a List<FinancialRecord> to WorkbookDesigner, processes the markers, and builds a pivot table that groups by Category and SubCategory with summed Amount. | Show how to programmatically add a new worksheet, set the source range from the populated smart‑marker area, and create a PivotTable using Aspose.Cells Pivot API. | Demonstrate refreshing and calculating the pivot table then saving the workbook as FinancialPivot.xlsx.
// Common Searches: how to generate a pivot table after processing smart markers with Aspose.Cells C# | asp.net example creating pivot from List of objects using WorkbookDesigner | c# Aspose.Cells populate worksheet with smart markers then add pivot table | financial data summary Excel pivot using Aspose.Cells programmatically
// Tags: Aspose.Cells smart markers data binding | Aspose.Cells create pivot table programmatically | C# generate Excel pivot from object list | financial summary pivot Aspose.Cells | WorkbookDesigner populate worksheet | PivotTable AddFieldToArea usage

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotAfterSmartMarkers
{
    // Simple POCO to hold financial data
    // The program creates a workbook template with smart markers, uses WorkbookDesigner to populate it from a List<FinancialRecord>, adds a new worksheet, builds a pivot table that rows by Category, columns by SubCategory, and sums Amount, refreshes and calculates the pivot, then saves the file as FinancialPivot.xlsx.
    public class FinancialRecord
    {
        public string Category { get; set; }
        public string SubCategory { get; set; }
        public double Amount { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook (template) and set up smart markers
            Workbook workbook = new Workbook();
            Worksheet templateSheet = workbook.Worksheets[0];
            templateSheet.Name = "Template";

            // Header row
            templateSheet.Cells["A1"].PutValue("Category");
            templateSheet.Cells["B1"].PutValue("SubCategory");
            templateSheet.Cells["C1"].PutValue("Amount");

            // Smart marker row – will be expanded by WorkbookDesigner
            templateSheet.Cells["A2"].PutValue("&=$Category");
            templateSheet.Cells["B2"].PutValue("&=$SubCategory");
            templateSheet.Cells["C2"].PutValue("&=$Amount");

            // Define the smart marker range (required name)
            templateSheet.Cells.CreateRange("A2:C2").Name = "_CellsSmartMarkers";

            // 2. Prepare sample financial data
            List<FinancialRecord> data = new List<FinancialRecord>
            {
                new FinancialRecord { Category = "Revenue", SubCategory = "Product A", Amount = 120000 },
                new FinancialRecord { Category = "Revenue", SubCategory = "Product B", Amount = 85000 },
                new FinancialRecord { Category = "Expense", SubCategory = "Salaries", Amount = 50000 },
                new FinancialRecord { Category = "Expense", SubCategory = "Marketing", Amount = 20000 },
                new FinancialRecord { Category = "Revenue", SubCategory = "Product C", Amount = 60000 },
                new FinancialRecord { Category = "Expense", SubCategory = "R&D", Amount = 30000 }
            };

            // 3. Process smart markers to populate the worksheet with the data
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("FinancialData", data);
            designer.Process(); // populates the range defined by _CellsSmartMarkers

            // 4. Determine the populated data range (including headers)
            // MaxDisplayRange gives the used range of the sheet
            string sourceRange = $"=Template!{templateSheet.Cells.MaxDisplayRange.Address}";

            // 5. Add a new worksheet that will host the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // 6. Create the pivot table using the source data range
            // Parameters: sourceData, destination cell (upper‑left corner), table name
            int pivotIndex = pivotSheet.PivotTables.Add(sourceRange, "A1", "FinancialPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // 7. Configure the pivot fields
            // Rows – Category, Columns – SubCategory, Data – Sum of Amount
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "SubCategory");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Optional: display the pivot in tabular form for better readability
            pivotTable.ShowInTabularForm();

            // 8. Refresh and calculate the pivot data so that values appear in the sheet
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // 9. Save the final workbook
            workbook.Save("FinancialPivot.xlsx");
        }
    }
}

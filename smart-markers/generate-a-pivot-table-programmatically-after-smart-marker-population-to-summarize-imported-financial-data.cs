using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and a worksheet that will hold the raw data
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Define column headers
        dataSheet.Cells["A1"].PutValue("Date");
        dataSheet.Cells["B1"].PutValue("Account");
        dataSheet.Cells["C1"].PutValue("Amount");

        // Insert smart markers – they will be replaced by the designer
        dataSheet.Cells["A2"].PutValue("&=$Date");
        dataSheet.Cells["B2"].PutValue("&=$Account");
        dataSheet.Cells["C2"].PutValue("&=$Amount");

        // Mark the range that contains the smart markers
        dataSheet.Cells.CreateRange("A2:C2").Name = "_CellsSmartMarkers";

        // Sample financial data that will populate the smart markers
        var records = new List<FinancialRecord>
        {
            new FinancialRecord { Date = new DateTime(2023, 1, 1), Account = "Revenue", Amount = 12000 },
            new FinancialRecord { Date = new DateTime(2023, 1, 2), Account = "Expense", Amount = 5000 },
            new FinancialRecord { Date = new DateTime(2023, 1, 3), Account = "Revenue", Amount = 8000 },
            new FinancialRecord { Date = new DateTime(2023, 1, 4), Account = "Expense", Amount = 3000 }
        };

        // Process smart markers and populate the worksheet with the data above
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource("FinancialData", records);
        designer.Process();

        // Add a new worksheet that will contain the pivot table
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

        // Build the source data reference for the pivot table (including headers)
        string sourceData = $"=Data!{dataSheet.Cells.MaxDisplayRange.Address}";

        // Add the pivot table to the pivot sheet (top‑left cell A1) and give it a name
        int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A1", "FinancialPivot");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

        // Configure the pivot fields:
        //   Row – Account
        //   Column – Date
        //   Data – Sum of Amount
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Account");
        pivotTable.AddFieldToArea(PivotFieldType.Column, "Date");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Optional layout and calculation
        pivotTable.ShowInTabularForm();
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the final workbook
        workbook.Save("FinancialPivotReport.xlsx");
    }

    // Simple POCO representing a financial record
    public class FinancialRecord
    {
        public DateTime Date { get; set; }
        public string Account { get; set; }
        public double Amount { get; set; }
    }
}
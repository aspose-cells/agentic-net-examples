using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class LoadOptionsParsingPivotCachedRecordsDemo
    {
        public static void Run()
        {
            // Create load options for XLSX format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            
            // Enable parsing of pivot cached records while loading
            loadOptions.ParsingPivotCachedRecords = true;
            
            // Load the workbook with the specified options
            Workbook workbook = new Workbook("input.xlsx", loadOptions);
            
            // Access the first worksheet and rename it
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "PivotData";
            
            // Add a pivot table to demonstrate that the option is effective
            // Source range: A1:C10, destination cell: D1, pivot table name: PivotTable1
            int pivotIndex = worksheet.PivotTables.Add("A1:C10", "D1", "PivotTable1");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
            
            // Add the first column as a row field and the second column as a data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);
            
            // Save the workbook with the new pivot table
            workbook.Save("PivotTableOutput.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LoadOptionsParsingPivotCachedRecordsDemo.Run();
        }
    }
}
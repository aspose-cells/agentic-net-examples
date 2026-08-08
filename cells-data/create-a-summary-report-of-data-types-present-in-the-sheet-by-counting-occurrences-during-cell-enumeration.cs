// Title: Create a CellValueType summary report in Excel with Aspose.Cells for .NET
// Description: A C# example that fills a workbook with numeric, string, DateTime, Boolean, blank and error values, exports the cell value types to a 2‑D array, counts each CellValueType, writes the totals to a new "Summary" worksheet, and saves the file as DataTypeSummary.xlsx.
// Keywords: Aspose.Cells | CellValueType | count cell types | export type array | C# Excel data profiling | summary worksheet | data type statistics | enumerate cells | Excel workbook analysis | Aspose.Cells .NET example
// Common Searches: how to count cell value types with Aspose.Cells | Aspose.Cells export type array C# | generate data type summary sheet Aspose.Cells | C# count numeric string date boolean cells in Excel | Aspose.Cells create summary worksheet programmatically
// Developer Intent: Enumerate every cell, tally each CellValueType, and produce a worksheet that lists the type names with their occurrence counts.
// Use Cases: Quick data‑quality audit showing the distribution of numbers, text, dates, booleans, blanks and errors. | Pre‑processing step to decide which transformations are needed based on cell type composition. | Add an automatic summary tab to generated reports for stakeholders to review data type breakdown.
// AI Prompts: Write C# code using Aspose.Cells that iterates over a worksheet, counts each CellValueType, and outputs the results to a new summary sheet. | Explain the ExportTypeArray method in Aspose.Cells and how to treat blank and error cells when summarizing types. | Suggest improvements for the summary sheet, such as sorting by count, adding percentage columns, or applying conditional formatting.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsDataTypeSummary
{
    // A C# example that fills a workbook with numeric, string, DateTime, Boolean, blank and error values, exports the cell value types to a 2‑D array, counts each CellValueType, writes the totals to a new "Summary" worksheet, and saves the file as DataTypeSummary.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data of various types
            sheet.Cells["A1"].PutValue(123);                     // Numeric
            sheet.Cells["B1"].PutValue("Hello World");           // String
            sheet.Cells["C1"].PutValue(DateTime.Now);            // DateTime
            sheet.Cells["D1"].PutValue(true);                    // Boolean
            sheet.Cells["E1"].PutValue(null);                    // Null (blank)
            sheet.Cells["A2"].PutValue(45.67);                   // Numeric
            sheet.Cells["B2"].PutValue("Aspose.Cells");          // String
            sheet.Cells["C2"].PutValue(false);                  // Boolean
            sheet.Cells["D2"].PutValue("");                     // String (empty)
            sheet.Cells["E2"].PutValue("#DIV/0!");               // Error (as string for demonstration)

            // Determine the used range dimensions
            int maxRow = sheet.Cells.MaxDataRow;
            int maxColumn = sheet.Cells.MaxDataColumn;
            int totalRows = maxRow + 1;      // rows are zero‑based
            int totalColumns = maxColumn + 1;

            // Export the cell value types to a 2‑D array
            CellValueType[,] typeArray = sheet.Cells.ExportTypeArray(0, 0, totalRows, totalColumns);

            // Count occurrences of each CellValueType
            Dictionary<CellValueType, int> typeCounts = new Dictionary<CellValueType, int>();
            for (int i = 0; i < totalRows; i++)
            {
                for (int j = 0; j < totalColumns; j++)
                {
                    CellValueType type = typeArray[i, j];
                    if (typeCounts.ContainsKey(type))
                        typeCounts[type]++;
                    else
                        typeCounts[type] = 1;
                }
            }

            // Add a new worksheet for the summary report
            Worksheet summarySheet = workbook.Worksheets.Add("Summary");
            summarySheet.Cells["A1"].PutValue("Cell Value Type");
            summarySheet.Cells["B1"].PutValue("Count");

            // Write the summary data
            int rowIndex = 1;
            foreach (var kvp in typeCounts)
            {
                summarySheet.Cells[rowIndex, 0].PutValue(kvp.Key.ToString());
                summarySheet.Cells[rowIndex, 1].PutValue(kvp.Value);
                rowIndex++;
            }

            // Save the workbook
            workbook.Save("DataTypeSummary.xlsx");
        }
    }
}

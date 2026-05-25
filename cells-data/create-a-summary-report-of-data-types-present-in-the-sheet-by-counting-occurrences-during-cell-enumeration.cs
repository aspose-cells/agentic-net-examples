using System;
using System.Collections.Generic;
using Aspose.Cells;

class DataTypeSummaryReport
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate the worksheet with sample data of various types
        worksheet.Cells["A1"].PutValue(123);                     // Numeric
        worksheet.Cells["A2"].PutValue("Hello World");          // String
        worksheet.Cells["A3"].PutValue(DateTime.Now);           // DateTime
        worksheet.Cells["A4"].PutValue(true);                  // Boolean
        worksheet.Cells["A5"].PutValue(null);                  // Null (blank)
        worksheet.Cells["B1"].PutValue(45.6);                   // Numeric
        worksheet.Cells["B2"].PutValue("Aspose.Cells");        // String
        // Cells C1:C3 are left empty to represent blanks

        // Dictionary to store counts of each CellValueType
        Dictionary<CellValueType, int> typeCounts = new Dictionary<CellValueType, int>();

        // Enumerate all instantiated cells in the worksheet
        foreach (Cell cell in worksheet.Cells)
        {
            CellValueType type = cell.Type;
            if (typeCounts.ContainsKey(type))
                typeCounts[type]++;
            else
                typeCounts[type] = 1;
        }

        // Add a new worksheet to hold the summary report
        Worksheet summarySheet = workbook.Worksheets.Add("Summary");
        int summaryRow = 0;

        // Write header
        summarySheet.Cells[summaryRow, 0].PutValue("Cell Value Type");
        summarySheet.Cells[summaryRow, 1].PutValue("Count");
        summaryRow++;

        // Write each type and its count
        foreach (KeyValuePair<CellValueType, int> entry in typeCounts)
        {
            summarySheet.Cells[summaryRow, 0].PutValue(entry.Key.ToString());
            summarySheet.Cells[summaryRow, 1].PutValue(entry.Value);
            summaryRow++;
        }

        // Save the workbook with the summary report
        workbook.Save("DataTypeSummaryReport.xlsx");
    }
}
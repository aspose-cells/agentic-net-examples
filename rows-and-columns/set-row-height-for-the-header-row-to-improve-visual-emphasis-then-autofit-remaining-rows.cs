// Title: Set a custom height for the header row and auto‑fit the remaining rows with Aspose.Cells for .NET (C#)
// AI Prompts: Set the first worksheet row to 30 points high and automatically size all subsequent rows using Aspose.Cells in C#. | Modify the example to use a different header height while preserving automatic row sizing only for data rows. | Insert code that auto‑size columns after adjusting row heights in an Aspose.Cells workbook.
// Common Searches: Aspose.Cells C# set header row height and auto size other rows | how to exclude first row from AutoFitRows in Aspose.Cells .NET | adjust row height in points then auto‑fit rows range Aspose.Cells | auto‑fit columns after setting custom row heights using Aspose.Cells C# | sample code for header row formatting with Aspose.Cells workbook
// Tags: header row custom height Aspose.Cells | row auto‑size excluding header .NET | set row height points C# | column width auto‑fit after row height Aspose.Cells | auto‑fit rows range without header Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Demonstrates creating a workbook, setting the first row height to 30 points, populating sample data, auto‑fitting rows 2‑n, auto‑fitting columns, and saving the file as XLSX.
class SetHeaderRowHeightAndAutoFit
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add header values
        cells["A1"].PutValue("Header 1");
        cells["B1"].PutValue("Header 2");
        cells["C1"].PutValue("Header 3");

        // Set a custom height for the header row (row index 0)
        worksheet.Cells.Rows[0].Height = 30; // height in points

        // Populate some sample data rows
        for (int i = 1; i <= 10; i++)
        {
            cells[i, 0].PutValue($"Item {i}");
            cells[i, 1].PutValue($"Description for item {i} that might be long and wrap");
            cells[i, 2].PutValue(i * 10);
        }

        // Auto‑fit all rows except the header row
        int startDataRow = 1;                       // first data row index
        int endDataRow = cells.MaxDataRow;          // last row that contains data
        worksheet.AutoFitRows(startDataRow, endDataRow);

        // Optional: auto‑fit columns for better visibility
        worksheet.AutoFitColumns();

        // Save the workbook
        string outputPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "HeaderRowHeightDemo.xlsx");
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}

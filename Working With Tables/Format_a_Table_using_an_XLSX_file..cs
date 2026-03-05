using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

class TableFormatter
{
    static void Main()
    {
        // Input and output file paths
        string inputFile = @"C:\Temp\InputWorkbook.xlsx";
        string outputFile = @"C:\Temp\FormattedWorkbook.xlsx";

        Workbook workbook;

        // Load existing workbook if it exists; otherwise create a new one with sample data and a table
        if (File.Exists(inputFile))
        {
            workbook = new Workbook(inputFile);
        }
        else
        {
            workbook = new Workbook();
            Worksheet ws = workbook.Worksheets[0];

            // Populate sample data
            ws.Cells["A1"].PutValue("ID");
            ws.Cells["B1"].PutValue("Name");
            ws.Cells["C1"].PutValue("Score");

            ws.Cells["A2"].PutValue(1);
            ws.Cells["B2"].PutValue("Alice");
            ws.Cells["C2"].PutValue(85);

            ws.Cells["A3"].PutValue(2);
            ws.Cells["B3"].PutValue("Bob");
            ws.Cells["C3"].PutValue(92);

            ws.Cells["A4"].PutValue(3);
            ws.Cells["B4"].PutValue("Charlie");
            ws.Cells["C4"].PutValue(78);

            // Create a table (ListObject) covering the data range
            int firstRow = 0; // zero‑based index
            int firstCol = 0;
            int totalRows = 4; // header + 3 data rows
            int totalCols = 3;
            ListObject table = ws.ListObjects[ws.ListObjects.Add(firstRow, firstCol, firstRow + totalRows - 1, firstCol + totalCols - 1, true)];
            table.DisplayName = "SampleTable";
        }

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one table
        if (worksheet.ListObjects.Count == 0)
        {
            Console.WriteLine("No tables found in the worksheet.");
            return;
        }

        // Retrieve the first table
        ListObject tableObj = worksheet.ListObjects[0];

        // Header row index (zero‑based)
        int headerRowIndex = tableObj.StartRow;

        // Create style for header row
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.Font.Color = Color.White;
        headerStyle.ForegroundColor = Color.DarkBlue;
        headerStyle.Pattern = BackgroundType.Solid;

        StyleFlag headerFlag = new StyleFlag { All = true };
        worksheet.Cells.ApplyRowStyle(headerRowIndex, headerStyle, headerFlag);

        // Apply alternating row style for data rows
        for (int row = tableObj.StartRow + 1; row <= tableObj.EndRow; row++)
        {
            if ((row - tableObj.StartRow) % 2 == 0) // even data row (zero‑based)
            {
                Style dataStyle = workbook.CreateStyle();
                dataStyle.ForegroundColor = Color.LightGray;
                dataStyle.Pattern = BackgroundType.Solid;

                StyleFlag dataFlag = new StyleFlag { All = true };
                worksheet.Cells.ApplyRowStyle(row, dataStyle, dataFlag);
            }
        }

        // Save the workbook
        workbook.Save(outputFile);
        Console.WriteLine($"Table formatted and saved to: {outputFile}");
    }
}
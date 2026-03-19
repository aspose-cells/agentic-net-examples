using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses Workbook() constructor)
            Workbook workbook = new Workbook();

            // Access the default worksheet and rename it
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data
            dataSheet.Cells["A1"].PutValue("Product");
            dataSheet.Cells["B1"].PutValue("Quantity");
            dataSheet.Cells["A2"].PutValue("Apple");
            dataSheet.Cells["B2"].PutValue(150);
            dataSheet.Cells["A3"].PutValue("Banana");
            dataSheet.Cells["B3"].PutValue(200);
            dataSheet.Cells["A4"].PutValue("Orange");
            dataSheet.Cells["B4"].PutValue(120);

            // Insert a new row at index 1 (second row) and shift existing rows down
            CellArea insertArea = CellArea.CreateCellArea(1, 0, 1, 1);
            dataSheet.Cells.InsertRange(insertArea, 1, ShiftType.Down, true);

            // Add headers for the inserted row
            dataSheet.Cells["A2"].PutValue("Category");
            dataSheet.Cells["B2"].PutValue("Type");

            // Save the workbook to a MemoryStream as an XLS file (uses SaveToStream)
            MemoryStream stream = workbook.SaveToStream();

            // Load a new workbook from the MemoryStream (uses Workbook(Stream) constructor)
            Workbook loadedWorkbook = new Workbook(stream);

            // Add a new worksheet for summary information
            int summaryIndex = loadedWorkbook.Worksheets.Add();
            Worksheet summarySheet = loadedWorkbook.Worksheets[summaryIndex];
            summarySheet.Name = "Summary";

            // Calculate total quantity from the data sheet
            int totalQuantity = 0;
            for (int row = 1; row <= 4; row++) // rows 2‑5 in Excel (0‑based index)
            {
                object cellValue = loadedWorkbook.Worksheets[0].Cells[row, 1].Value;
                if (cellValue != null && int.TryParse(cellValue.ToString(), out int qty))
                {
                    totalQuantity += qty;
                }
            }

            // Write summary data
            summarySheet.Cells["A1"].PutValue("Total Quantity");
            summarySheet.Cells["B1"].PutValue(totalQuantity);

            // Save the final workbook as XLSX (uses Workbook.Save(string))
            loadedWorkbook.Save("DemoOutput.xlsx");

            // Dispose resources
            stream.Dispose();
            workbook.Dispose();
            loadedWorkbook.Dispose();

            Console.WriteLine("Workbook created and saved as DemoOutput.xlsx");
        }
    }
}
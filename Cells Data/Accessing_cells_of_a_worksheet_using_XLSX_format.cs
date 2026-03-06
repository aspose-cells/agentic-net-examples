using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsAccessCellsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (XLSX format is the default)
            Workbook workbook = new Workbook();

            // Get the first worksheet in the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the Cells collection from the worksheet
            Cells cells = worksheet.Cells;

            // -------------------------------------------------
            // Access cells using numeric indexes (row, column)
            // Row and column indexes are zero‑based.
            // -------------------------------------------------
            cells[0, 0].PutValue("Product"); // Cell A1
            cells[0, 1].PutValue("Price");   // Cell B1

            // -------------------------------------------------
            // Access cells using the A1 style notation
            // -------------------------------------------------
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(1.25);
            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(0.80);

            // -------------------------------------------------
            // Apply a simple style to the header row (A1:B1)
            // -------------------------------------------------
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;

            cells[0, 0].SetStyle(headerStyle);
            cells[0, 1].SetStyle(headerStyle);

            // -------------------------------------------------
            // Save the workbook as an XLSX file
            // -------------------------------------------------
            workbook.Save("Products.xlsx", SaveFormat.Xlsx);
        }
    }
}
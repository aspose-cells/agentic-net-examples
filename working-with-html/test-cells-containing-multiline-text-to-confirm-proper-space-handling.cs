using System;
using Aspose.Cells;

namespace AsposeCellsMultiLineCellTest
{
    // Author: Aspose.Cells .NET example – testing multi‑line text handling in cells
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define a cell that contains multi‑line text using '\n'
            Cell multiLineCell = sheet.Cells["A1"];
            multiLineCell.Value = "First line\nSecond line\nThird line";

            // Enable text wrapping so the newlines are respected when displayed
            Style wrapStyle = multiLineCell.GetStyle();
            wrapStyle.IsTextWrapped = true;
            multiLineCell.SetStyle(wrapStyle);

            // Auto‑fit the row height to accommodate the wrapped text
            sheet.AutoFitRow(0);

            // Save the workbook (save rule)
            string filePath = "MultiLineCellTest.xlsx";
            workbook.Save(filePath);

            // Reload the workbook to verify persistence (load rule)
            Workbook loadedWorkbook = new Workbook(filePath);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            Cell loadedCell = loadedSheet.Cells["A1"];

            // Verify that the cell still contains the newline characters
            string cellText = loadedCell.StringValue;
            Console.WriteLine("Cell text contains {0} lines.", cellText.Split('\n').Length);

            // Verify that text wrapping is still enabled
            bool isWrapped = loadedCell.GetStyle().IsTextWrapped;
            Console.WriteLine("Text wrapping enabled: " + isWrapped);
        }
    }
}
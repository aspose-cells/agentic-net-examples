using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook from disk
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Modify built‑in document properties
            workbook.BuiltInDocumentProperties["Author"].Value = "Jane Doe";
            workbook.BuiltInDocumentProperties["Title"].Value = "Modified Workbook";

            // Add a custom document property
            workbook.CustomDocumentProperties.Add("Reviewed", true);

            // Change the default style (font name and size)
            workbook.DefaultStyle.Font.Name = "Calibri";
            workbook.DefaultStyle.Font.Size = 11;

            // Add a new worksheet and insert some data
            int newSheetIndex = workbook.Worksheets.Add();
            Worksheet newSheet = workbook.Worksheets[newSheetIndex];
            newSheet.Name = "Summary";
            newSheet.Cells["A1"].PutValue("Report generated on:");
            newSheet.Cells["B1"].PutValue(DateTime.Now);

            // Save the modified workbook as XLSX, preserving all changes
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}
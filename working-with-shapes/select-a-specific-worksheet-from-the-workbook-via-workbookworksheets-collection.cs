using System;
using Aspose.Cells;

namespace WorksheetSelectionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Rename the default worksheet
            Worksheet firstSheet = workbook.Worksheets[0];
            firstSheet.Name = "DataSheet";

            // Add a second worksheet and set its code name
            Worksheet secondSheet = workbook.Worksheets.Add("ReportSheet");
            secondSheet.CodeName = "Report";

            // Select a worksheet by index (zero‑based)
            Worksheet sheetByIndex = workbook.Worksheets[0];
            sheetByIndex.Cells["A1"].PutValue("Selected by index");

            // Select a worksheet by its name
            Worksheet sheetByName = workbook.Worksheets["ReportSheet"];
            sheetByName.Cells["A1"].PutValue("Selected by name");

            // Select a worksheet by its code name
            Worksheet sheetByCodeName = workbook.Worksheets.GetSheetByCodeName("Report");
            sheetByCodeName.Cells["A2"].PutValue("Selected by code name");

            // Save the workbook
            workbook.Save("WorksheetSelectionDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
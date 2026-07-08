using System;
using Aspose.Cells;

namespace AsposeCellsSelectWorksheet
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle start)
            Workbook workbook = new Workbook();

            // Rename the default first worksheet for demonstration
            Worksheet firstSheet = workbook.Worksheets[0];
            firstSheet.Name = "DataSheet";

            // Add another worksheet
            Worksheet secondSheet = workbook.Worksheets.Add("ReportSheet");

            // Select a worksheet by its name
            // Here we select the worksheet we just renamed to "DataSheet"
            Worksheet selectedSheet = workbook.Worksheets["DataSheet"];

            // Obtain the Cells collection from the selected worksheet
            Cells cells = selectedSheet.Cells;

            // Perform some operations on the cells collection
            cells["A1"].PutValue("Hello");
            cells["B2"].PutValue(12345);
            cells["C3"].Formula = "=SUM(B2,B2)";

            // Calculate formulas to ensure the result is up‑to‑date
            workbook.CalculateFormula();

            // Output a cell value to the console for verification
            Console.WriteLine("C3 value (SUM of B2+B2): " + cells["C3"].IntValue);

            // Save the workbook (lifecycle end)
            workbook.Save("SelectWorksheetDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
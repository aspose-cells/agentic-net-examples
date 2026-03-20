using System;
using Aspose.Cells;

namespace ConvertStringToNumericDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook and add sample data ----------
            Workbook wbCreate = new Workbook();                     // create workbook
            Cells cellsCreate = wbCreate.Worksheets[0].Cells;       // get cells collection

            // Put string values that represent numbers and a non‑numeric string
            cellsCreate["A1"].PutValue("123.45");
            cellsCreate["A2"].PutValue("678");
            cellsCreate["A3"].PutValue("NotANumber");

            // Convert all convertible string values to proper numeric types
            cellsCreate.ConvertStringToNumericValue();              // use Cells.ConvertStringToNumericValue method

            // Save the workbook that was created and converted
            wbCreate.Save("CreatedConverted.xlsx");                 // save workbook

            // ---------- Load an existing workbook, convert, and save ----------
            Workbook wbLoad = new Workbook("input.xlsx");            // load workbook from file
            Cells cellsLoad = wbLoad.Worksheets[0].Cells;           // get cells collection

            // Perform the same conversion on the loaded workbook
            cellsLoad.ConvertStringToNumericValue();                // convert strings to numeric where possible

            // Save the modified workbook
            wbLoad.Save("LoadedConverted.xlsx");                    // save workbook
        }
    }
}
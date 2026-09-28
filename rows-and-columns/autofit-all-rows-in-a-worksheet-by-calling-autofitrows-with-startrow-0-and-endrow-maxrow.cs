// Title: How to auto‑fit every row in an Aspose.Cells worksheet from the first to the last populated row using C#
// AI Prompts: Write C# code that creates a new Aspose.Cells workbook, adds sample data, determines the last occupied row with Cells.MaxDataRow, and calls the AutoFitRows method from row 0 through that index. | Provide a step‑by‑step example of using Aspose.Cells in .NET to automatically resize the height of all rows by supplying the first row and the maximum data row to the auto‑fit API.
// Common Searches: Aspose.Cells C# auto fit rows from first to last data row | Worksheet.AutoFitRows startRow 0 endRow MaxDataRow sample code | C# adjust row height for entire worksheet using Aspose.Cells | determine last populated row in Aspose.Cells before auto‑fitting rows
// Tags: Aspose.Cells auto‑fit rows range | C# MaxDataRow row height adjustment | auto‑fit all worksheet rows .NET | adjust Excel row height Aspose.Cells

using System;
using Aspose.Cells;

namespace AutoFitRowsExample
{
    // The program creates a new workbook, populates a few cells with varying text, retrieves the zero‑based index of the last row that contains data via Cells.MaxDataRow, auto‑fits rows from index 0 to that row using AutoFitRows, and saves the file as AutoFitRowsResult.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data so that rows have content to autofit
            worksheet.Cells["A1"].PutValue("Short text");
            worksheet.Cells["A2"].PutValue("This is a longer piece of text that should cause the row height to increase when autofitted.");
            worksheet.Cells["A3"].PutValue("Another line with\nmultiple line breaks\nto test row height adjustment.");

            // Determine the last row that contains data
            int maxRow = worksheet.Cells.MaxDataRow; // zero‑based index of the last used row

            // Auto‑fit all rows from the first (0) to the last used row
            worksheet.AutoFitRows(0, maxRow);

            // Save the workbook to a file
            workbook.Save("AutoFitRowsResult.xlsx", SaveFormat.Xlsx);
        }
    }
}

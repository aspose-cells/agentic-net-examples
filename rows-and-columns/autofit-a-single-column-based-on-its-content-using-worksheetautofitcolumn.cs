using System;
using Aspose.Cells;

namespace AutoFitColumnExample
{
    // Author: Aspose.Cells .NET example
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some cells in column C (index 2)
            worksheet.Cells["C1"].PutValue("Short");
            worksheet.Cells["C2"].PutValue("A much longer piece of text that requires wider column");
            worksheet.Cells["C3"].PutValue("Medium length");

            // Auto‑fit column C (0‑based index 2)
            worksheet.AutoFitColumn(2);

            // Save the workbook to a file
            workbook.Save("AutoFitColumnResult.xlsx");
        }
    }
}
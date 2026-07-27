using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data in the first row
        worksheet.Cells["A1"].PutValue("Short");
        worksheet.Cells["B1"].PutValue("A much longer piece of text that may require a taller row.");

        // Define the row index and the desired minimum height (in points)
        int rowIndex = 0;
        double minHeight = 20.0; // Minimum row height

        // Set the minimum height first
        worksheet.Cells.SetRowHeight(rowIndex, minHeight);

        // Auto‑fit the row based on its content
        worksheet.AutoFitRow(rowIndex);

        // Ensure the row height is not less than the minimum
        double actualHeight = worksheet.Cells.GetRowHeight(rowIndex);
        if (actualHeight < minHeight)
        {
            worksheet.Cells.SetRowHeight(rowIndex, minHeight);
        }

        // Save the workbook
        workbook.Save("RowHeight_MinAutoFit.xlsx");
    }
}
// Author: Aspose.Cells .NET example – combines SetRowHeight with AutoFitRow to enforce a minimum row height.
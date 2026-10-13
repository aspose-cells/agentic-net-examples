// Title: Add a top‑most average subtotal row to column F with Aspose.Cells for .NET
// AI Prompts: Create a C# workbook, insert a new first row, and assign an AVERAGE formula that computes the mean of the existing values in column F using Aspose.Cells. | Programmatically determine the data range in column F, place a subtotal row at row 1, and set the cell formula to AVERAGE(F2:Fn) with Aspose.Cells. | Generate an Excel file where the average of column F is displayed in a header row above the data, leveraging the Aspose.Cells Worksheet.InsertRows and Cells.Formula properties.
// Common Searches: how to insert a header row with an average formula for column F using Aspose.Cells in C# | Aspose.Cells calculate column average and place result at the top of the sheet .NET | C# code to add a subtotal row before data in an Excel worksheet with Aspose.Cells
// Tags: insert top row Aspose.Cells .NET | set AVERAGE formula column F Aspose.Cells | worksheet subtotal row Excel Aspose.Cells | calculate column average programmatically Aspose.Cells | save workbook output.xlsx Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, fills column F with sample values, inserts a row at the very top, writes an AVERAGE formula in the first cell of column F that references the data range, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // (Optional) Populate column F with sample data for demonstration
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 5].PutValue(i + 1); // Column F has zero‑based index 5
            }

            // Insert a new row at the top to hold the average (subtotal) value
            sheet.Cells.InsertRows(0, 1);

            // Determine the range that contains the data (excluding the newly inserted row)
            int firstDataRow = 1; // after insertion, data starts at row index 1
            int lastDataRow = sheet.Cells.MaxDataRow;

            // Set a formula in the first cell of column F to calculate the average of the data range
            sheet.Cells[0, 5].Formula = $"AVERAGE(F{firstDataRow + 1}:F{lastDataRow + 1})";

            // Save the workbook to a file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

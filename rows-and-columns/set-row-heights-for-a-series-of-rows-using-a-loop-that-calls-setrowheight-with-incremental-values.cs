// Title: How to set incremental row heights for multiple rows in an Excel sheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook with Aspose.Cells, loops through rows 0‑9, and sets each row height starting at 15 pt and increasing by 2 pt per row, then saves the file as .xlsx. | Generate a .NET example that demonstrates using Cells.SetRowHeight inside a for‑loop to apply varying heights to consecutive rows.
// Common Searches: Aspose.Cells C# set different row heights in a loop | programmatically increase Excel row height by 2 points per row using Aspose | example of using SetRowHeight for multiple rows with Aspose.Cells .NET | how to apply incremental row heights to rows 1‑10 in an Excel file with C# | loop through rows and assign custom heights with Aspose.Cells API
// Tags: Aspose.Cells SetRowHeight C# loop | incremental row height Aspose.Cells | dynamic Excel row height .NET | for-loop row height automation Aspose | Excel worksheet row height programmatic

using System;
using Aspose.Cells;

namespace AsposeCellsRowHeightDemo
{
    // Creates a new workbook, iterates rows 0‑9, sets each row height starting at 15 pt and adding 2 pt per row using Cells.SetRowHeight, and saves the result as RowHeightsDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Loop through rows 0 to 9 and set incremental heights
            // Height starts at 15 points and increases by 2 points for each subsequent row
            for (int rowIndex = 0; rowIndex < 10; rowIndex++)
            {
                double height = 15.0 + (rowIndex * 2.0);
                cells.SetRowHeight(rowIndex, height);
            }

            // Save the workbook to a file
            workbook.Save("RowHeightsDemo.xlsx");
        }
    }
}

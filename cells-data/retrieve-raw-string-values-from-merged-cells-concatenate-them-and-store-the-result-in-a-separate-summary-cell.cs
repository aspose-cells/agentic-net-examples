// Title: Concatenate raw string values from merged cells and store the result in a summary cell with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to loop through all merged cell areas, read each cell's StringValue, concatenate the texts, and write the combined string into a specified summary cell. | Show an example of extracting raw string content from merged ranges in a worksheet, building a single summary string, and saving the workbook with the summary placed in cell E1 using Aspose.Cells.
// Common Searches: Aspose.Cells C# concatenate values from merged cell ranges into one cell | how to read StringValue of each cell in merged areas with Aspose.Cells | C# example for summarizing merged cell contents in Excel using Aspose.Cells | loop through merged cells and build a summary string Aspose.Cells .NET | store concatenated merged cell text in a separate cell with Aspose.Cells
// Tags: concatenate merged cell stringvalues Aspose.Cells | retrieve raw string from merged ranges .NET | write summary cell after merging Aspose.Cells | iterate merged cell areas C# | StringValue extraction merged cells Aspose.Cells | merged cells aggregation Excel .NET

using System;
using System.Text;
using Aspose.Cells;

namespace MergedCellsSummary
{
    // The program creates a workbook, merges two ranges (A1:B2 and A3:B4), iterates over each merged area, reads every cell's raw StringValue, concatenates the texts, writes the combined result into cell E1, and saves the file as MergedCellsSummary.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Sample merged cells for demonstration
            // Merge A1:B2 and put a value
            worksheet.Cells.Merge(0, 0, 2, 2);
            worksheet.Cells[0, 0].PutValue("First");

            // Merge A3:B4 and put a value
            worksheet.Cells.Merge(2, 0, 2, 2);
            worksheet.Cells[2, 0].PutValue("Second");

            // Retrieve all merged areas
            CellArea[] mergedAreas = worksheet.Cells.GetMergedAreas();

            // Concatenate raw string values from each merged cell
            StringBuilder summaryBuilder = new StringBuilder();

            foreach (CellArea area in mergedAreas)
            {
                for (int row = area.StartRow; row <= area.EndRow; row++)
                {
                    for (int col = area.StartColumn; col <= area.EndColumn; col++)
                    {
                        Cell cell = worksheet.Cells[row, col];
                        if (cell != null && cell.Value != null)
                        {
                            // Use StringValue to get the raw string representation
                            summaryBuilder.Append(cell.StringValue);
                        }
                    }
                }
            }

            // Store the concatenated result in a separate summary cell (e.g., E1)
            worksheet.Cells["E1"].PutValue(summaryBuilder.ToString());

            // Save the workbook
            workbook.Save("MergedCellsSummary.xlsx");
        }
    }
}

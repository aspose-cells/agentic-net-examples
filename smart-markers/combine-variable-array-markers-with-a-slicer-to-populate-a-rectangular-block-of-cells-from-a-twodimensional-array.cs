// Title: Use Aspose.Cells Smart Markers to Fill a 2x2 Excel range from a sliced jagged int array in C#
// AI Prompts: Write C# code that extracts rows 1‑2 and columns 1‑2 from a jagged int[][] using the C# range operator, converts the 2x2 slice into a DataTable, places variable array smart markers (&=Block, &=Block.Col0, &=Block.Col1) in a worksheet, processes them with WorkbookDesigner, and saves the workbook as an .xlsx file. | Show how to map a sliced two‑dimensional array to Aspose.Cells smart markers so each element populates a rectangular block of cells, including dynamic DataTable schema creation and WorkbookDesigner invocation.
// Common Searches: aspocells c# slice jagged array and fill excel range with smart markers | using variable array markers to write sub‑array data into an .xlsx file with Aspose.Cells | transform sliced int array into a DataTable for Aspose.Cells WorkbookDesigner | example of C# range slicer feeding smart markers in Aspose.Cells | fill a 2x2 cell range from a jagged array slice using Aspose.Cells
// Tags: smart markers slice jagged array to datatable | c# range operator aspocells workbookdesigner | variable array markers rectangular block | int[][] slice conversion to datatable aspocells | excel block filling via smart markers c#

using System;
using System.Data;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program creates a Workbook, defines a 4x4 jagged int array, slices rows 1‑2 and columns 1‑2 with C# range syntax to obtain a 2x2 sub‑array, converts that slice into a DataTable, inserts variable array smart markers (&=Block, &=Block.Col0, &=Block.Col1) into cells A1‑B2, processes the markers with WorkbookDesigner, and saves the result as Output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define a two‑dimensional jagged array
                int[][] sourceArray = new int[][]
                {
                    new int[] { 1, 2, 3, 4 },
                    new int[] { 5, 6, 7, 8 },
                    new int[] { 9, 10, 11, 12 },
                    new int[] { 13, 14, 15, 16 }
                };

                // Use C# range slicer to obtain a rectangular sub‑array (rows 1‑2, columns 1‑2)
                // Note: slicing a jagged array requires slicing each inner array separately
                int[][] slicedArray = sourceArray[1..3]                     // rows 1 and 2 (0‑based)
                    .Select(row => row[1..3].ToArray())                    // columns 1 and 2
                    .ToArray();                                            // result is 2x2 array

                // Convert the sliced jagged array into a DataTable – the format expected by Smart Markers
                DataTable dt = new DataTable("Block");

                // Add columns dynamically based on the inner array length
                for (int col = 0; col < slicedArray[0].Length; col++)
                {
                    dt.Columns.Add("Col" + col, typeof(int));
                }

                // Populate rows
                foreach (int[] row in slicedArray)
                {
                    dt.Rows.Add(row.Cast<object>().ToArray());
                }

                // Place variable array markers in the worksheet.
                // The marker "&=Block" tells Aspose.Cells to repeat the row for each DataTable row.
                // The markers "&=Block.Col0", "&=Block.Col1", … refer to individual column values.
                sheet.Cells["A1"].PutValue("&=Block");
                sheet.Cells["A2"].PutValue("&=Block.Col0");
                sheet.Cells["B2"].PutValue("&=Block.Col1");

                // Process the smart markers using WorkbookDesigner
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dt);
                designer.Process();

                // Save the workbook (lifecycle rule: save)
                workbook.Save("Output.xlsx");
                Console.WriteLine("Workbook saved successfully as Output.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

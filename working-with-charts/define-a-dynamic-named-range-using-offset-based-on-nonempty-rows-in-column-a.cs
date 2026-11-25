// Title: Define a dynamic named range with OFFSET and COUNTA for column A using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# create dynamic named range that expands with non‑empty rows in column A | How to use OFFSET and COUNTA in Aspose.Cells to define an automatically growing range | Programmatically add a named range based on the length of column A data with Aspose.Cells .NET | Retrieve address and size of a dynamic range after adding OFFSET formula in Aspose.Cells | Set a value in the last cell of a dynamic named range using Aspose.Cells C#
// Tags: OFFSET formula dynamic range Aspose.Cells | COUNTA based range expansion .NET | add named range programmatically Aspose.Cells | retrieve range by name C# Aspose.Cells | save workbook with dynamic named range Aspose.Cells

using System;
using Aspose.Cells;

namespace DynamicNamedRangeDemo
{
    // The example creates a new workbook, populates column A with a header and items, and defines a dynamic named range called "DynamicRange" using an OFFSET formula that leverages COUNTA to count non‑empty cells in column A. The range is added to the worksheet's Names collection, retrieved to display its address, row count, and column count, a value is written to the last cell of the range, and the workbook is saved as DynamicRangeDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                string sheetName = sheet.Name; // Usually "Sheet1"

                // Populate column A with sample data (including a header)
                sheet.Cells["A1"].PutValue("Header");
                for (int i = 2; i <= 10; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                }

                // Define a dynamic named range that starts at A2 and expands
                // downwards based on the number of non‑empty cells in column A.
                // Formula: =OFFSET(Sheet1!$A$2,0,0,COUNTA(Sheet1!$A:$A)-1,1)
                string formula = $"=OFFSET('{sheetName}'!$A$2,0,0,COUNTA('{sheetName}'!$A:$A)-1,1)";

                // Add the named range to the workbook's name collection
                int nameIndex = workbook.Worksheets.Names.Add("DynamicRange");
                Name dynamicName = workbook.Worksheets.Names[nameIndex];
                dynamicName.RefersTo = formula;

                // Retrieve the range by its name to demonstrate that it works
                Aspose.Cells.Range dynamicRange = workbook.Worksheets.GetRangeByName("DynamicRange");
                Console.WriteLine($"DynamicRange address: {dynamicRange.Address}");
                Console.WriteLine($"Rows in DynamicRange: {dynamicRange.RowCount}");
                Console.WriteLine($"Columns in DynamicRange: {dynamicRange.ColumnCount}");

                // Example usage: set a value in the last cell of the dynamic range
                int lastRow = dynamicRange.FirstRow + dynamicRange.RowCount - 1;
                int col = dynamicRange.FirstColumn;
                sheet.Cells[lastRow, col].PutValue("LastItem");

                // Save the workbook
                string outputPath = "DynamicRangeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}

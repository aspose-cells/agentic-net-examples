// Title: Delete a column from an Excel ListObject and remove its formulas using Aspose.Cells for .NET (C#)
// Common Searches: C# code to delete a column from an Excel ListObject using Aspose.Cells | How to remove a table column and its formulas with Aspose.Cells for .NET | Aspose.Cells ListColumns.RemoveAt example for clearing column data | Removing a column from an Excel table without leaving formula remnants in .NET | Delete specific ListObject column by index with Aspose.Cells
// Tags: ListObject column removal Aspose.Cells C# | delete Excel table column formulas Aspose.Cells | ListColumns.RemoveAt usage .NET | modify ListObject structure without residual formulas | Aspose.Cells table column management

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The sample loads a workbook, accesses the first worksheet's ListObject, validates a zero‑based column index, removes the specified column with ListColumns.RemoveAt (which also deletes any formulas in that column), and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook containing a ListObject (Excel table)
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index or name as needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one ListObject
                if (sheet.ListObjects.Count == 0)
                {
                    Console.WriteLine("No ListObject (table) found in the worksheet.");
                    return;
                }

                // Retrieve the first ListObject on the worksheet
                ListObject listObject = sheet.ListObjects[0];

                // Specify the zero‑based index of the column to remove (e.g., the third column)
                int columnIndexToRemove = 2;

                // Validate column index
                if (columnIndexToRemove < 0 || columnIndexToRemove >= listObject.ListColumns.Count)
                {
                    Console.WriteLine("Column index to remove is out of range.");
                    return;
                }

                // Remove the column from the ListObject; this also removes any formulas in that column
                listObject.ListColumns.RemoveAt(columnIndexToRemove);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

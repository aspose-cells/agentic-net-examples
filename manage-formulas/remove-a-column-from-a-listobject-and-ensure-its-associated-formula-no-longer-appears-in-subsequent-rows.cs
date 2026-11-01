// Title: Delete a ListObject column and eliminate its formula with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# remove a column from an Excel table without keeping the formula | How to clear ListColumn formulas before deleting the column in Aspose.Cells | C# code to delete the 'Amount' column from a ListObject and stop formula propagation | Removing a table column and its associated formula using Aspose.Cells .NET API | Aspose.Cells example for deleting ListObject column and clearing formulas
// Tags: delete ListObject column Aspose.Cells | clear ListColumn formula C# | remove Excel table column Aspose.Cells .NET | prevent formula propagation Aspose.Cells | Aspose.Cells ListObject column management

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The sample loads an existing workbook, checks for a ListObject on the first worksheet, finds the column named "Amount", clears any formula attached to that ListColumn, removes the column from the table, and saves the result to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "Input.xlsx";
            string outputPath = "Output.xlsx";

            try
            {
                // Verify input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one table (ListObject)
                if (sheet.ListObjects.Count == 0)
                {
                    Console.WriteLine("No tables (ListObjects) found in the worksheet.");
                    return;
                }

                // Get the first table on the worksheet
                ListObject table = sheet.ListObjects[0];

                // Retrieve the column named "Amount"
                ListColumn column = table.ListColumns["Amount"];

                if (column != null)
                {
                    // Clear any formula to prevent propagation (if applicable)
                    try
                    {
                        column.Formula = string.Empty;
                    }
                    catch
                    {
                        // Some versions may not support Formula property; ignore if not available
                    }

                    // Determine the column's index within the table
                    int columnIndex = table.ListColumns.IndexOf(column);

                    // Remove the column from the table
                    table.ListColumns.RemoveAt(columnIndex);
                    Console.WriteLine("Column 'Amount' removed successfully.");
                }
                else
                {
                    Console.WriteLine("Column 'Amount' not found in the table.");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display the message
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

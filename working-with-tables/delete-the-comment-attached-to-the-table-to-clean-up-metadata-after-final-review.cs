// Title: Aspose.Cells for .NET – Remove a Table (ListObject) Comment from an Excel Workbook
// Description: C# example that loads an Excel file with Aspose.Cells, checks for tables (ListObjects), clears the Comment property of the first table, and saves the workbook without the table comment. Includes error handling for missing files and empty worksheets.
// Keywords: Aspose.Cells remove table comment | clear ListObject comment C# | delete Excel table metadata Aspose | Aspose.Cells ListObject Comment property | C# remove Excel table comment
// Common Searches: how to delete a table comment using Aspose.Cells .NET | Aspose.Cells clear ListObject comment before saving | remove Excel table comment C# Aspose example | set table comment to empty Aspose.Cells
// Developer Intent: Delete the comment attached to a ListObject (table) in an Excel workbook and save the updated file.
// Use Cases: Programmatically clean up table metadata by clearing the Comment field of a specific table. | Validate that a worksheet contains at least one table before attempting to modify its comment. | Provide user feedback when no tables are present or when the input file cannot be found.
// AI Prompts: Generate C# code using Aspose.Cells that removes comments from all tables in every worksheet, with robust file‑existence checks. | Create a reusable method that iterates through a workbook's worksheets, clears each ListObject's Comment property, and returns the modified workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // C# example that loads an Excel file with Aspose.Cells, checks for tables (ListObjects), clears the Comment property of the first table, and saves the workbook without the table comment. Includes error handling for missing files and empty worksheets.
    public class DeleteTableComment
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "InputWithTable.xlsx";
            const string outputPath = "OutputWithoutTableComment.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook that contains the table (list object)
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet has at least one table
                if (worksheet.ListObjects.Count > 0)
                {
                    // Get the first table (ListObject)
                    ListObject table = worksheet.ListObjects[0];

                    // Delete the comment attached to the table by clearing the Comment property
                    table.Comment = string.Empty;
                }
                else
                {
                    Console.WriteLine("No tables found in the worksheet.");
                }

                // Save the workbook after removing the table comment
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}

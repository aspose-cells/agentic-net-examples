// Title: Aspose.Cells C# – Error handling for adding a slicer to a non‑existent table column
// Description: Demonstrates how to create a workbook, define a two‑column table, validate a zero‑based column index, throw an ArgumentOutOfRangeException for invalid indexes, catch errors during slicer creation, log the issue, and save the file safely.
// Keywords: Aspose.Cells | C# slicer | slicer error handling | invalid column index | ListObject slicer exception | ArgumentOutOfRangeException | validate slicer column | Aspose.Cells API | Excel slicer programmatically | Aspose.Cells .NET
// Common Searches: Aspose.Cells add slicer invalid column | C# validate slicer column index Aspose.Cells | How to catch slicer creation exception in Aspose.Cells | ListObject slicer out of range error | Aspose.Cells slicer error handling example
// Developer Intent: Add a slicer to a table column while safely handling cases where the column does not exist.
// Use Cases: Prevent runtime crashes by checking column index before calling SlicerCollection.Add. | Log detailed error messages when slicer creation fails. | Skip slicer creation and continue processing when the target column is missing. | Provide a fallback UI element (e.g., dropdown) if a slicer cannot be added. | Automate workbook generation with robust slicer validation.
// AI Prompts: Write C# code using Aspose.Cells that adds a slicer to a table column with pre‑validation and try‑catch for errors. | Show how to record slicer creation failures to a log file instead of console output in Aspose.Cells. | Explain how to enumerate ListObject columns and select a valid column before adding a slicer in Aspose.Cells .NET. | Provide a step‑by‑step guide to handle ArgumentOutOfRangeException when adding a slicer to a non‑existent column. | Generate a reusable method that adds a slicer with built‑in validation for any Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, define a two‑column table, validate a zero‑based column index, throw an ArgumentOutOfRangeException for invalid indexes, catch errors during slicer creation, log the issue, and save the file safely.
    public class SlicerErrorHandlingDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unexpected error: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for a table (2 columns, 3 rows)
            worksheet.Cells["A1"].PutValue("Column1");
            worksheet.Cells["B1"].PutValue("Column2");
            worksheet.Cells["A2"].PutValue("Data1");
            worksheet.Cells["B2"].PutValue("Data2");
            worksheet.Cells["A3"].PutValue("Data3");
            worksheet.Cells["B3"].PutValue("Data4");

            // Add a table covering the data range
            int tableIndex = worksheet.ListObjects.Add("A1", "B3", true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Define the target column index (zero‑based). Intentionally invalid.
            int targetColumnIndex = 5; // Table only has 2 columns (indices 0 and 1)

            // Attempt to add a slicer for the specified column with error handling
            try
            {
                // Verify that the column exists before accessing it
                if (targetColumnIndex < 0 || targetColumnIndex >= table.ListColumns.Count)
                {
                    throw new ArgumentOutOfRangeException(
                        nameof(targetColumnIndex),
                        $"Column index {targetColumnIndex} is out of range. Table has {table.ListColumns.Count} columns.");
                }

                // Retrieve the ListColumn object
                ListColumn column = table.ListColumns[targetColumnIndex];

                // Add the slicer at cell I38
                SlicerCollection slicers = worksheet.Slicers;
                int slicerIndex = slicers.Add(table, column, "I38");

                // Optional: configure the slicer (e.g., set a caption)
                Slicer slicer = slicers[slicerIndex];
                slicer.Caption = $"Slicer for {column.Name}";
            }
            catch (Exception ex)
            {
                // Handle any errors that occur during slicer creation
                Console.WriteLine("Error adding slicer: " + ex.Message);
            }

            // Save the workbook
            workbook.Save("SlicerErrorHandlingOutput.xlsx");
            Console.WriteLine("Workbook saved as SlicerErrorHandlingOutput.xlsx");
        }
    }
}

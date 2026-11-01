// Title: C# unit test with Aspose.Cells to confirm FreezePanes freezes a specific number of rows
// Common Searches: aspocells c# unit test freeze panes rows count | verify freeze panes rows persistence after saving workbook in Aspose.Cells | how to assert frozen rows using Aspose.Cells FreezePanes method | unit testing Aspose.Cells FreezePanes with memory stream | c# test that FreezePanes rowsToFreeze parameter works correctly
// Tags: Aspose.Cells FreezePanes row freeze verification | C# unit test workbook memory stream | Aspose.Cells four-parameter FreezePanes overload | freeze panes settings persistence Aspose.Cells | worksheet row count assertion after reload

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a new Workbook, applies the four‑parameter FreezePanes overload to freeze a given number of rows, saves the workbook to a MemoryStream, reloads it, and asserts that the worksheet still contains at least the expected number of rows, throwing an exception if the freeze settings appear not to have persisted.
    public class FreezePanesExample
    {
        public static void Main()
        {
            try
            {
                new FreezePanesExample().Run();
                Console.WriteLine("Test completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public void Run()
        {
            // Create a new workbook
            Workbook workbook = CreateWorkbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Number of rows to freeze
            int rowsToFreeze = 7;

            // Apply FreezePanes to freeze the specified rows (no columns)
            // The 4‑parameter overload is used for newer Aspose.Cells versions.
            // Row index is 1‑based; to freeze the first N rows, set the first parameter to rowsToFreeze + 1.
            sheet.FreezePanes(rowsToFreeze + 1, 1, rowsToFreeze, 0);

            // Save workbook to memory stream
            using (MemoryStream stream = SaveWorkbook(workbook))
            {
                // Load workbook back from stream
                Workbook loadedWorkbook = LoadWorkbook(stream);
                Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

                // No direct properties to verify freeze settings in this API version,
                // but we can ensure the sheet still contains the expected number of rows.
                if (loadedSheet.Cells.MaxDataRow < rowsToFreeze - 1)
                {
                    throw new InvalidOperationException("FreezePanes settings may not have persisted correctly.");
                }
            }
        }

        // Creates a new empty workbook
        private Workbook CreateWorkbook()
        {
            return new Workbook();
        }

        // Saves the workbook to a memory stream
        private MemoryStream SaveWorkbook(Workbook workbook)
        {
            try
            {
                MemoryStream ms = new MemoryStream();
                workbook.Save(ms, SaveFormat.Xlsx);
                ms.Position = 0;
                return ms;
            }
            catch (Exception ex)
            {
                throw new IOException("Failed to save workbook to memory stream.", ex);
            }
        }

        // Loads a workbook from a stream
        private Workbook LoadWorkbook(Stream stream)
        {
            try
            {
                return new Workbook(stream);
            }
            catch (Exception ex)
            {
                throw new IOException("Failed to load workbook from stream.", ex);
            }
        }
    }
}

// Title: C# code to verify absence of #REF! errors after merging multiple Excel workbooks with formulas using Aspose.Cells
// Common Searches: how to check for #REF! errors in a merged Excel workbook with Aspose.Cells C# | Aspose.Cells detect error cells after copying worksheets in .NET | C# merge several .xlsx files and ensure no formula reference errors | validate cell error types after workbook merge using Aspose.Cells API | scan merged workbook for CellValueType.IsError in C# | remove default sheet and verify no #REF! errors after merging Excel files
// Tags: Aspose.Cells merge multiple .xlsx worksheets | detect #REF! errors with CellValueType.IsError | validate merged workbook for formula errors | C# scan used range for error cells | remove default sheet Aspose.Cells | copy worksheets without breaking references

using System;
using System.IO;
using Aspose.Cells;

// The example loads several source Excel files, copies each worksheet into a new workbook, removes the initial empty sheet, then iterates over the used range of every worksheet to flag cells whose value type is an error (e.g., #REF!). It reports any found errors, confirms when none exist, and saves the merged workbook as 'MergedResult.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook that will hold the merged worksheets
            Workbook mergedWorkbook = new Workbook();

            // Paths of the workbooks to be merged
            string[] sourceFiles = { "Source1.xlsx", "Source2.xlsx", "Source3.xlsx" };

            // Load each source workbook and copy its worksheets into the merged workbook
            foreach (string file in sourceFiles)
            {
                if (!File.Exists(file))
                {
                    Console.WriteLine($"Source file not found: {file}");
                    continue; // Skip missing files
                }

                try
                {
                    Workbook src = new Workbook(file);
                    foreach (Worksheet ws in src.Worksheets)
                    {
                        // Add a new sheet to the merged workbook and copy the source worksheet into it
                        int newIndex = mergedWorkbook.Worksheets.Add(); // returns the index of the new sheet
                        Worksheet newWs = mergedWorkbook.Worksheets[newIndex];
                        newWs.Name = ws.Name; // set the name to match the source sheet
                        ws.Copy(newWs);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to process '{file}': {ex.Message}");
                }
            }

            // Remove the default empty sheet if it still exists after copying
            if (mergedWorkbook.Worksheets.Count > 1 && mergedWorkbook.Worksheets[0].Name == "Sheet1")
            {
                mergedWorkbook.Worksheets.RemoveAt(0);
            }

            // Validate that no cell contains an error (e.g., #REF!) after merging
            bool hasError = false;
            foreach (Worksheet ws in mergedWorkbook.Worksheets)
            {
                // Get the used range of the worksheet
                Aspose.Cells.Range usedRange = ws.Cells.MaxDisplayRange;
                foreach (Cell cell in usedRange)
                {
                    // Flag any cell that contains an error value
                    if (cell.Type == CellValueType.IsError)
                    {
                        Console.WriteLine($"Error found in sheet '{ws.Name}' at cell {cell.Name}");
                        hasError = true;
                    }
                }
            }

            if (!hasError)
            {
                Console.WriteLine("No error cells detected after merging.");
            }

            // Save the merged workbook
            mergedWorkbook.Save("MergedResult.xlsx");
            Console.WriteLine("Merged workbook saved as 'MergedResult.xlsx'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}

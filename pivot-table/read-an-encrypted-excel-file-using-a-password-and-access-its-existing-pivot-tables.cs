// Title: C# – Open Password‑Protected Excel File and List Pivot Tables with Aspose.Cells
// Description: Shows how to load an encrypted .xlsx using Aspose.Cells LoadOptions with a password, fallback to an unprotected load, iterate worksheets, access each PivotTableCollection, refresh and calculate the pivots, and output pivot names and row‑field counts.
// Keywords: Aspose.Cells | C# password protected Excel | load encrypted workbook | pivot table enumeration | RefreshData | CalculateData | .NET Excel API | open encrypted .xlsx | pivot tables programmatically | Excel security
// Common Searches: Aspose.Cells open encrypted Excel C# | list pivot tables in protected workbook Aspose | refresh pivot tables after loading password Excel | C# read .xlsx with password and get pivot tables | how to use LoadOptions password Aspose.Cells
// Developer Intent: Load a password‑protected workbook and retrieve its pivot tables programmatically.
// Use Cases: Securely open a workbook that requires a password. | Automatically retry opening without a password when the supplied one is invalid. | Enumerate all pivot tables across worksheets for reporting or validation. | Refresh and recalculate pivot data to reflect the latest source values. | Log pivot table names and row‑field counts for audit purposes.
// AI Prompts: Generate C# code using Aspose.Cells to open an encrypted .xlsx with a given password and handle incorrect passwords gracefully. | Write a function that returns a dictionary of pivot table names and their row‑field counts from a loaded workbook. | Explain the steps to refresh and calculate pivot tables after loading a password‑protected workbook with Aspose.Cells. | Provide a sample that iterates worksheets and prints pivot table details in C#.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExample
{
    // Shows how to load an encrypted .xlsx using Aspose.Cells LoadOptions with a password, fallback to an unprotected load, iterate worksheets, access each PivotTableCollection, refresh and calculate the pivots, and output pivot names and row‑field counts.
    class Program
    {
        static void Main()
        {
            // Path to the encrypted Excel file
            string filePath = "encrypted.xlsx";

            // Password used to protect the workbook (if any)
            string password = "myPassword";

            // Verify that the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File \"{filePath}\" not found.");
                return;
            }

            Workbook workbook = null;

            // Attempt to open the workbook with the supplied password
            try
            {
                LoadOptions loadOptions = new LoadOptions { Password = password };
                workbook = new Workbook(filePath, loadOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to open with password: {ex.Message}");
                Console.WriteLine("Attempting to open without a password...");

                // Try opening without a password
                try
                {
                    workbook = new Workbook(filePath);
                }
                catch (Exception innerEx)
                {
                    Console.WriteLine($"Failed to open workbook: {innerEx.Message}");
                    return;
                }
            }

            // Ensure workbook was loaded
            if (workbook == null)
            {
                Console.WriteLine("Workbook could not be loaded.");
                return;
            }

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet: {sheet.Name}");

                // Access the collection of pivot tables in the current worksheet
                PivotTableCollection pivotTables = sheet.PivotTables;

                // Loop through each pivot table
                for (int i = 0; i < pivotTables.Count; i++)
                {
                    PivotTable pivot = pivotTables[i];
                    Console.WriteLine($"  PivotTable {i}: {pivot.Name}");

                    try
                    {
                        // Refresh the pivot table data (optional)
                        pivot.RefreshData();
                        pivot.CalculateData();

                        // Output the number of row fields in the pivot table
                        Console.WriteLine($"    Row fields count: {pivot.RowFields.Count}");
                    }
                    catch (Exception pivotEx)
                    {
                        Console.WriteLine($"    Error processing pivot table: {pivotEx.Message}");
                    }
                }
            }
        }
    }
}

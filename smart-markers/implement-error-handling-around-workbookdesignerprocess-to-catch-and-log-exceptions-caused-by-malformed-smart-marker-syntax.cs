// Title: C# error handling for WorkbookDesigner.Process – catch malformed smart marker exceptions in Aspose.Cells
// Description: Shows how to enclose WorkbookDesigner.Process in a try‑catch block, capture Aspose.Cells CellsException caused by invalid smart marker syntax, log the message and exception code, and still save the workbook.
// Keywords: Aspose.Cells | WorkbookDesigner | smart markers | error handling | exception handling | CellsException | malformed smart marker | C# .NET | try catch | logging example | GitHub sample
// Common Searches: how to handle smart marker errors in Aspose.Cells | catch CellsException when processing smart markers C# | log malformed smart marker exception Aspose.Cells | WorkbookDesigner.Process try catch example | Aspose.Cells smart marker syntax validation
// Developer Intent: Wrap WorkbookDesigner.Process in a try‑catch to detect and log errors from invalid smart marker syntax.
// Use Cases: Prevent application crash by catching CellsException from malformed smart markers. | Record detailed error information (message and code) for troubleshooting. | Continue workbook generation and saving even when smart marker processing fails.
// AI Prompts: Generate C# code that surrounds WorkbookDesigner.Process with try‑catch, logs CellsException details, and saves the workbook. | Show how to replace console output with a file logger for Aspose.Cells smart marker errors. | Provide a pre‑validation routine that checks smart marker syntax before calling Process to avoid exceptions.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerErrorHandling
{
    // Shows how to enclose WorkbookDesigner.Process in a try‑catch block, capture Aspose.Cells CellsException caused by invalid smart marker syntax, log the message and exception code, and still save the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a malformed smart marker (incorrect syntax)
            // Correct syntax would be like "&=Table.Column"
            // The following marker is intentionally malformed to trigger an exception
            sheet.Cells["A1"].PutValue("&=Invalid..Marker");

            // Prepare a simple data source
            DataTable dt = new DataTable("ValidTable");
            dt.Columns.Add("Column1", typeof(string));
            dt.Rows.Add("Value1");

            // Initialize WorkbookDesigner and assign the workbook (lifecycle rule)
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Set the data source
            designer.SetDataSource(dt);

            try
            {
                // Process the smart markers (operation we want to protect)
                designer.Process();
            }
            catch (CellsException ex)
            {
                // Log detailed information about Aspose.Cells specific exceptions
                Console.WriteLine("Aspose.Cells exception occurred while processing smart markers.");
                Console.WriteLine($"Message: {ex.Message}");
                Console.WriteLine($"Exception Type Code: {ex.Code}");
            }
            catch (Exception ex)
            {
                // Log any other unexpected exceptions
                Console.WriteLine("An unexpected error occurred while processing smart markers.");
                Console.WriteLine($"Message: {ex.Message}");
            }

            // Save the workbook (saving rule)
            workbook.Save("ProcessedWorkbook.xlsx");
        }
    }
}

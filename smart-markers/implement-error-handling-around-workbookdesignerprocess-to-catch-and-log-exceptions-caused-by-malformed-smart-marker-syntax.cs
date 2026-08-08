// Title: C# – Handle malformed smart marker errors with WorkbookDesigner.Process in Aspose.Cells
// Description: Shows how to wrap WorkbookDesigner.Process in try‑catch blocks to capture CellsException and generic errors caused by invalid smart marker syntax, log details, and still save the workbook.
// Keywords: Aspose.Cells | WorkbookDesigner.Process | smart marker error handling | CellsException | C# exception handling | invalid smart marker syntax | log Aspose.Cells errors | save workbook after failure
// Common Searches: catch CellsException when processing smart markers Aspose.Cells | handle malformed smart marker syntax C# | log errors from WorkbookDesigner.Process | save workbook after smart marker processing error | Aspose.Cells smart marker try catch example
// Developer Intent: Wrap WorkbookDesigner.Process in try‑catch to detect and log errors from incorrect smart marker syntax while ensuring the workbook can still be saved.
// Use Cases: Detect and log CellsException when a smart marker is missing the '=' character. | Continue program execution and save the original workbook even if processing fails. | Expose the Aspose.Cells exception code (ex.Code) for troubleshooting smart marker issues.
// AI Prompts: Generate C# code that adds detailed logging (including stack trace) around WorkbookDesigner.Process for smart marker errors. | Create a reusable method that processes smart markers with exception handling and returns a success flag. | Show how to integrate NLog (or another logging framework) with Aspose.Cells smart marker error handling in C#.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerErrorHandling
{
    // Shows how to wrap WorkbookDesigner.Process in try‑catch blocks to capture CellsException and generic errors caused by invalid smart marker syntax, log details, and still save the workbook.
    public class SmartMarkerProcessor
    {
        public static void Run()
        {
            // Create a new workbook (template)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a malformed smart marker (incorrect syntax)
            // Correct syntax would be like "&=Table.Column", here we omit the "="
            sheet.Cells["A1"].PutValue("&Table.Column");

            // Set up a simple data source so that processing can be attempted
            DataTable dt = new DataTable("Table");
            dt.Columns.Add("Column", typeof(string));
            dt.Rows.Add("Sample Value");

            // Initialize the WorkbookDesigner and assign the workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Bind the data source
            designer.SetDataSource(dt);

            // Attempt to process the smart markers with error handling
            try
            {
                designer.Process();
                Console.WriteLine("Smart markers processed successfully.");
            }
            catch (CellsException ex) // Specific Aspose.Cells exception
            {
                Console.WriteLine("Aspose.Cells exception occurred during processing:");
                Console.WriteLine($"Message: {ex.Message}");
                Console.WriteLine($"Exception Type Code: {ex.Code}");
            }
            catch (Exception ex) // General exception fallback
            {
                Console.WriteLine("An unexpected error occurred during processing:");
                Console.WriteLine($"Message: {ex.Message}");
            }

            // Save the workbook (even if processing failed, the file will contain the original content)
            try
            {
                workbook.Save("ProcessedWorkbook.xlsx");
                Console.WriteLine("Workbook saved as 'ProcessedWorkbook.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                SmartMarkerProcessor.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}

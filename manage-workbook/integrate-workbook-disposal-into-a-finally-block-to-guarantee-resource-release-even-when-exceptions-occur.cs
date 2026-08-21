// Title: C# – Ensure Aspose.Cells Workbook disposal with a finally clause
// Description: This example declares a Workbook outside a try block, writes data to cells, saves the file, catches any errors, and reliably releases unmanaged resources by calling workbook?.Dispose() inside a finally section.
// Keywords: Aspose.Cells | C# workbook disposal | finally clause | resource cleanup | Dispose pattern | exception safe | unmanaged resources
// Common Searches: Aspose.Cells dispose in C# | C# try finally workbook cleanup | release Aspose.Cells resources after save | how to ensure workbook is disposed on error | best practice for Aspose.Cells memory management
// Developer Intent: Implement guaranteed disposal of the Aspose.Cells Workbook regardless of exceptions.
// Use Cases: Generate a spreadsheet, add data, and guarantee memory is freed even if Save fails. | Integrate safe workbook handling into batch jobs or web services. | Replace using statements with explicit finally disposal for legacy code.
// AI Prompts: Write C# code that creates an Aspose.Cells Workbook, adds values, saves to 'output.xlsx', and disposes it in a finally block with null checking. | Show how to refactor Aspose.Cells Workbook usage from a using statement to a try‑catch‑finally pattern for explicit disposal. | Explain why calling workbook?.Dispose() in finally is preferred when the workbook variable is declared outside the try block.

using System;
using Aspose.Cells;

// This example declares a Workbook outside a try block, writes data to cells, saves the file, catches any errors, and reliably releases unmanaged resources by calling workbook?.Dispose() inside a finally section.
class WorkbookDisposeExample
{
    static void Main()
    {
        Workbook workbook = null; // Declare outside try to access in finally
        try
        {
            // Create a new workbook instance (uses the provided constructor rule)
            workbook = new Workbook();

            // Access the first worksheet and add some data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Save the workbook to disk (uses the provided Save method rule)
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during processing
            Console.WriteLine($"Error: {ex.Message}");
        }
        finally
        {
            // Guarantee that the workbook releases unmanaged resources
            workbook?.Dispose();
        }
    }
}

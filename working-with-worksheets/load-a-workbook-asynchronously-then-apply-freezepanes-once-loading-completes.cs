// Title: Async Load Excel Workbook and Apply FreezePanes with Aspose.Cells (C#)
// Description: Demonstrates how to load an Excel file on a background thread using `await Task.Run`, freeze the first three rows and columns at cell C3 with `Worksheet.FreezePanes`, and save the updated workbook—all without blocking the UI.
// Keywords: Aspose.Cells | C# async workbook loading | FreezePanes | Task.Run Excel | background thread Excel processing | load large Excel file without UI freeze | Excel freeze panes C3 | asynchronous Excel manipulation | .NET Excel API
// Common Searches: Aspose.Cells load workbook asynchronously | C# freeze panes after async load | How to use FreezePanes with async Aspose.Cells | Prevent UI blocking when opening Excel in .NET | Async Excel file processing Aspose
// Developer Intent: Load an Excel workbook asynchronously, apply FreezePanes at cell C3, and save the result.
// Use Cases: Responsive desktop or web apps that need to open large Excel files without freezing the UI. | Batch processing pipelines that load many workbooks in parallel and set freeze panes for consistent navigation. | Automated report generators that freeze header rows/columns for better readability in the final Excel output.
// AI Prompts: Generate code to freeze panes at a cell address calculated from worksheet dimensions. | Show how to add ConfigureAwait(false) to the async workbook load for library usage in ASP.NET. | Provide error‑handling patterns for async workbook loading and FreezePanes execution.

using System;
using System.Threading.Tasks;
using Aspose.Cells;

// Demonstrates how to load an Excel file on a background thread using `await Task.Run`, freeze the first three rows and columns at cell C3 with `Worksheet.FreezePanes`, and save the updated workbook—all without blocking the UI.
class Program
{
    static async Task Main()
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";

        // Path where the modified workbook will be saved
        string outputPath = "output.xlsx";

        // Load the workbook on a background thread to avoid blocking the UI thread
        Workbook workbook = await Task.Run(() => new Workbook(inputPath));

        // Get the first worksheet in the workbook
        Worksheet worksheet = workbook.Worksheets[0];

        // Freeze panes at cell C3 with 3 frozen rows and 3 frozen columns
        worksheet.FreezePanes("C3", 3, 3);

        // Save the workbook with the applied freeze panes
        workbook.Save(outputPath);
    }
}

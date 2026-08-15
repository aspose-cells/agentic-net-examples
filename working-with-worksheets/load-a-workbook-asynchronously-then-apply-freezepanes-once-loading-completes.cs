// Title: Async Load Excel Workbook and Apply Freeze Panes with Aspose.Cells for .NET
// Description: Shows how to load an Excel file on a background thread using Aspose.Cells, set freeze panes at cell C3 (3 rows × 3 columns), and save the workbook within an async C# method.
// Keywords: Aspose.Cells async loading | C# freeze panes | Excel workbook Task.Run | non‑blocking Excel processing | apply FreezePanes C3 | asynchronous Excel manipulation .NET | background thread workbook load | Aspose.Cells .NET example | freeze panes Excel C# | async Excel save
// Common Searches: aspacells load workbook asynchronously | c# async freeze panes aspocells | how to use FreezePanes after async load | aspocells async example for large Excel files | non blocking excel processing aspocells
// Developer Intent: Load an Excel workbook without blocking, apply freeze panes at C3, and persist the changes.
// Use Cases: Web API endpoints that generate Excel reports while keeping request threads responsive. | Desktop utilities that process large spreadsheets in the UI thread without freezing the interface. | Background services that prepare Excel files for downstream systems, needing frozen header rows/columns.
// AI Prompts: Create an async Aspose.Cells routine that loads a workbook, freezes panes at D5 with 2 rows and 2 columns, and returns the workbook object. | Add comprehensive error handling and cancellation support to an async method that loads a workbook and applies FreezePanes using Aspose.Cells. | Show how to chain multiple asynchronous Aspose.Cells operations—load, set FreezePanes, format cells, and save—while preserving thread safety.

using System;
using System.Threading.Tasks;
using Aspose.Cells;

// Shows how to load an Excel file on a background thread using Aspose.Cells, set freeze panes at cell C3 (3 rows × 3 columns), and save the workbook within an async C# method.
class Program
{
    // Asynchronously loads a workbook, applies FreezePanes, and saves it.
    static async Task ApplyFreezePanesAsync(string inputFile, string outputFile)
    {
        // Load the workbook on a background thread to avoid blocking.
        Workbook workbook = await Task.Run(() => new Workbook(inputFile));

        // Access the first worksheet.
        Worksheet sheet = workbook.Worksheets[0];

        // Freeze panes at cell C3 with 3 frozen rows and 3 frozen columns.
        sheet.FreezePanes("C3", 3, 3);

        // Save the modified workbook.
        workbook.Save(outputFile);
    }

    static async Task Main(string[] args)
    {
        // Example usage: provide input and output file paths.
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        await ApplyFreezePanesAsync(inputPath, outputPath);

        Console.WriteLine("Workbook processed and saved to " + outputPath);
    }
}

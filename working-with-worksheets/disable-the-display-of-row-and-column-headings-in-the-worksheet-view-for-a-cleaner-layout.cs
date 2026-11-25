// Title: Hide row and column headers in an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide worksheet row and column headers | set IsRowColumnHeadersVisible false Aspose.Cells example | remove Excel row/column headings programmatically using Aspose.Cells | how to disable header display in an Aspose.Cells workbook | C# code to hide Excel sheet headers with Aspose.Cells
// Tags: IsRowColumnHeadersVisible Aspose.Cells | worksheet header visibility .NET | Aspose.Cells view options C# | Excel sheet header suppression Aspose | programmatic Excel UI customization Aspose.Cells

using System;
using Aspose.Cells;

// The sample loads or creates a workbook, accesses the first worksheet, sets IsRowColumnHeadersVisible to false to hide row and column headings, adds a couple of cells for content, and saves the file as output.xlsx.
class Program
{
    static void Main(string[] args)
    {
        // Output file name (can be changed as needed)
        const string outputFile = "output.xlsx";

        // Optional: first argument can be an existing workbook to modify
        string inputFile = args.Length > 0 ? args[0] : null;

        try
        {
            Workbook workbook;

            if (!string.IsNullOrEmpty(inputFile))
            {
                // Load an existing workbook
                workbook = new Workbook(inputFile);
                Console.WriteLine($"Loaded workbook from \"{inputFile}\".");
            }
            else
            {
                // Create a brand‑new workbook
                workbook = new Workbook();
                Console.WriteLine("Created a new workbook.");
            }

            // Access the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Disable the display of row and column headings
            sheet.IsRowColumnHeadersVisible = false;
            Console.WriteLine("Row and column headers visibility set to false.");

            // Add a couple of cells so the saved file has visible content
            sheet.Cells["A1"].PutValue("Headers are hidden");
            sheet.Cells["B2"].PutValue(12345);

            // Save the workbook
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            // Report any errors without swallowing them
            Console.Error.WriteLine("An error occurred:");
            Console.Error.WriteLine(ex.Message);
            Console.Error.WriteLine(ex.StackTrace);
        }
    }
}

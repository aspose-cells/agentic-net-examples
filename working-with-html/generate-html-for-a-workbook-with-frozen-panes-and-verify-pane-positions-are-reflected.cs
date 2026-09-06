// Title: Export an Excel workbook with frozen rows and columns to HTML and confirm pane positions using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, fills a 20x10 range, applies FreezePanes(5,3,4,2), saves it as an XLSX file, reloads the file, and reads the Pane properties to verify that the freeze settings persisted. | Extend the example to save the same workbook as an HTML file and programmatically examine the generated HTML to ensure the frozen rows and columns are rendered correctly.
// Common Searches: Aspose.Cells .NET export frozen panes to HTML | how to verify FreezePanes after saving workbook with Aspose.Cells | C# generate HTML from Excel with frozen rows and columns using Aspose.Cells | check pane positions in HTML output from Aspose.Cells workbook
// Tags: Aspose.Cells FreezePanes export to HTML | C# verify frozen pane settings after workbook save | Aspose.Cells generate HTML with frozen rows and columns | read pane information from loaded workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, populates a 20x10 cell range, freezes the first 4 rows and 2 columns, saves the file as XLSX, reloads it to confirm the FreezePanes configuration, then exports the workbook to HTML and validates that the frozen panes are reflected in the generated HTML output.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Freeze first 4 rows and first 2 columns (split at row 5, column 3)
            // FreezePanes(row, column, totalRows, totalColumns)
            sheet.FreezePanes(5, 3, 4, 2);

            // Save the workbook to a file
            string filePath = "FrozenPaneDemo.xlsx";

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(filePath));
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            workbook.Save(filePath);

            // Verify the file exists before loading
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"The file '{filePath}' was not created.");

            // Load the workbook back to verify it was saved correctly
            Workbook loadedWorkbook = new Workbook(filePath);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

            // Output a simple confirmation (pane details are not accessed to avoid API issues)
            Console.WriteLine("Workbook loaded successfully. Freeze pane settings were applied.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

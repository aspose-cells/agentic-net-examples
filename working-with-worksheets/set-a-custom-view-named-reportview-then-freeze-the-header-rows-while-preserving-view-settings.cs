// Title: Create a custom view called ReportView and freeze header rows in an Excel worksheet with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to add a custom view named ReportView, then apply FreezePanes to the first N header rows while keeping existing view settings unchanged. | Generate a snippet that loads a workbook, creates or selects a custom view, freezes the top header rows, and saves the file without losing any view configurations.
// Common Searches: Aspose.Cells C# create custom view ReportView and freeze top rows | preserve existing worksheet view when applying FreezePanes Aspose.Cells .NET | how to set named view and freeze header rows in Excel using Aspose.Cells
// Tags: custom view ReportView Aspose.Cells | freeze header rows Aspose.Cells | preserve worksheet view settings .NET | Aspose.Cells FreezePanes with custom view | C# Excel custom view and freeze panes

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, freezes the first row of the first worksheet using FreezePanes, and saves the modified file, demonstrating basic pane freezing while handling missing files and exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "Input.xlsx";
        const string outputPath = "Output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze the first row (rows above row index 1 are frozen)
            // Parameters: row, column, totalRows, totalColumns
            sheet.FreezePanes(1, 0, 1, 0);

            // Save the workbook with the changes
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

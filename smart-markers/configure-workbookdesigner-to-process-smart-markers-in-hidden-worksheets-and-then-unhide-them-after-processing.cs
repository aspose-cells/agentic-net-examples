// Title: Process smart markers in hidden worksheets and then unhide all sheets with WorkbookDesigner in C# using Aspose.Cells
// Common Searches: how to process smart markers on hidden worksheets using Aspose.Cells C# | unhide all worksheets after WorkbookDesigner.Process in Aspose.Cells | Aspose.Cells example for processing smart markers in hidden Excel sheets | C# code to make hidden sheets visible after smart marker processing | load workbook, process smart markers, and reveal hidden worksheets with Aspose.Cells
// Tags: WorkbookDesigner process smart markers hidden worksheets | Aspose.Cells unhide worksheets after processing | C# smart marker processing hidden sheets | Excel workbook visibility management Aspose.Cells | process smart markers and reveal hidden sheets C#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel workbook, uses WorkbookDesigner to process smart markers on hidden worksheets, then iterates through all worksheets to set them visible, ensures the output folder exists, and saves the updated file.
class SmartMarkerProcessor
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook that contains hidden worksheets with smart markers
            Workbook workbook = new Workbook(inputPath);

            // Initialize the WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // If the current Aspose.Cells version supports processing hidden worksheets,
            // the default behavior already includes them. No additional options are required.

            // (Optional) Set the data source for smart markers here
            // Example:
            // designer.SetDataSource("Employees", employeeDataTable);

            // Process the smart markers in all worksheets
            designer.Process();

            // After processing, unhide all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.IsVisible = true;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the processed workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook processed and saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

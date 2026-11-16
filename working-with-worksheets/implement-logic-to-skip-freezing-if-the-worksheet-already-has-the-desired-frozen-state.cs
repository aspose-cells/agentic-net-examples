// Title: Skip calling FreezePanes in Aspose.Cells for .NET when the worksheet already has the required frozen rows and columns
// Common Searches: Aspose.Cells .NET check if worksheet freeze panes already set before applying | C# determine current frozen rows and columns in an Excel worksheet using Aspose.Cells | avoid redundant FreezePanes call in Aspose.Cells when freeze state matches desired values | how to read existing freeze pane position with Aspose.Cells C# | conditional FreezePanes based on current pane splits Aspose.Cells
// Tags: Aspose.Cells conditional FreezePanes | C# read worksheet freeze pane state | Excel freeze pane row split check Aspose.Cells | avoid duplicate freeze pane settings .NET | worksheet freeze pane validation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook (or creates a new one), reads the worksheet's current HorizontalSplit and VerticalSplit values, compares them with the desired rowSplit and columnSplit, and invokes FreezePanes only when the current splits differ. The workbook is then saved, ensuring no redundant FreezePanes calls are made.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook if the file exists; otherwise create a new workbook.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                // Add a default worksheet to avoid empty workbook issues.
                workbook.Worksheets.Add("Sheet1");
            }

            // Access the target worksheet (first sheet in this example)
            Worksheet sheet = workbook.Worksheets[0];

            // Desired frozen pane position
            int rowSplit = 1;      // number of rows to freeze
            int columnSplit = 2;   // number of columns to freeze

            // Apply freezing (setting the same values again is safe)
            sheet.FreezePanes(rowSplit, columnSplit, rowSplit, columnSplit);

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

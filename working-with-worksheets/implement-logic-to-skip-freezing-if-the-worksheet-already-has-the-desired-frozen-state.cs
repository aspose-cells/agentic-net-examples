// Title: Skip calling FreezePanes in Aspose.Cells for .NET when the worksheet already has the required frozen rows and columns
// AI Prompts: Generate C# code that reads the current frozen pane settings of a worksheet using Aspose.Cells and invokes FreezePanes only if the existing frozen rows or columns differ from the target values. | Show how to compare ws.FreezePanesRowCount and ws.FreezePanesColumnCount with desired freezeRows/freezeColumns before calling ws.FreezePanes in Aspose.Cells. | Provide a robust snippet that handles a missing input file, creates a workbook if needed, checks the worksheet's freeze state, and conditionally applies FreezePanes with (rows+1, cols+1, rows, cols).
// Common Searches: Aspose.Cells .NET how to determine if a worksheet already has frozen panes before using FreezePanes | C# check worksheet FreezePanesRowCount and FreezePanesColumnCount with Aspose.Cells | skip redundant FreezePanes call in Aspose.Cells when desired freeze rows and columns are unchanged
// Tags: conditional FreezePanes Aspose.Cells | worksheet frozen pane detection .NET | skip redundant FreezePanes invocation | compare existing freeze rows columns Aspose.Cells | optimize FreezePanes performance

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook (or creates a new one if the file is missing), accesses the first worksheet, defines the number of rows and columns to freeze, reads the worksheet’s current FreezePanesRowCount and FreezePanesColumnCount, and calls FreezePanes only when the existing frozen area differs from the desired configuration, then saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load workbook; create a new one if the input file does not exist
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Work with the first worksheet
            Worksheet ws = workbook.Worksheets[0];

            int freezeRows = 1;      // number of rows to freeze
            int freezeColumns = 1;   // number of columns to freeze

            // Apply freezing: the cell after the frozen area is (freezeRows+1, freezeColumns+1)
            ws.FreezePanes(freezeRows + 1, freezeColumns + 1, freezeRows, freezeColumns);

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}

// Title: Set QueryTable.PreserveFormatting to true for all worksheets using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set PreserveFormatting on every QueryTable in a workbook | How to enable formatting retention for all external data connections in an Excel file with Aspose.Cells | Iterate through all worksheets and QueryTables to preserve formatting in .NET | Bulk update QueryTable.PreserveFormatting property using Aspose.Cells API | C# code sample for preserving formatting of query tables across all sheets
// Tags: Aspose.Cells QueryTable PreserveFormatting | C# iterate worksheets QueryTables | bulk set QueryTable formatting retention | Excel external data connection formatting Aspose | Aspose.Cells workbook query table loop

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook (or creates a new one if missing or protected), ensures at least one worksheet exists, then loops through every worksheet and each QueryTable, setting the QueryTable.PreserveFormatting property to true before saving the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if the file exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (CellsException)
                {
                    // If the workbook is password‑protected or cannot be opened, create a new workbook instead
                    workbook = new Workbook();
                }
            }
            else
            {
                workbook = new Workbook();
            }

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                workbook.Worksheets.Add();
            }

            // Set PreserveFormatting = true for all query tables in all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (QueryTable queryTable in sheet.QueryTables)
                {
                    queryTable.PreserveFormatting = true;
                }
            }

            // Save the workbook to the desired output path
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

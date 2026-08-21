// Title: Disable QueryTable PreserveFormatting in Aspose.Cells for .NET
// Description: Demonstrates how to locate a QueryTable in a worksheet and set its PreserveFormatting property to false, ensuring that each data refresh uses the workbook's default cell styles before saving the file.
// Keywords: Aspose.Cells QueryTable PreserveFormatting | disable PreserveFormatting .NET | default formatting on query refresh | Aspose.Cells C# QueryTable settings | turn off formatting preservation
// Common Searches: Aspose.Cells turn off PreserveFormatting | QueryTable default formatting after refresh | C# disable PreserveFormatting in Aspose.Cells | how to stop query table from keeping styles | Aspose.Cells refresh query without preserving format
// Developer Intent: Programmatically turn off the PreserveFormatting flag of an existing QueryTable so that refreshed data inherits the worksheet's current formatting.
// Use Cases: Prepare a workbook for periodic data imports where the visual style must stay consistent. | Automate report generation that pulls external data without retaining source cell formats. | Update legacy spreadsheets that contain QueryTables to respect new theme or style changes.
// AI Prompts: Generate C# code using Aspose.Cells to set PreserveFormatting = false for all QueryTables in a workbook. | Explain the impact of the PreserveFormatting property on QueryTable refresh behavior and how to modify it. | Show how to refresh a QueryTable after disabling PreserveFormatting and verify that default cell styles are applied.

using System;
using Aspose.Cells;

// Demonstrates how to locate a QueryTable in a worksheet and set its PreserveFormatting property to false, ensuring that each data refresh uses the workbook's default cell styles before saving the file.
class DisableQueryTablePreserveFormatting
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample data (simulating the source of a query table)
        worksheet.Cells["A1"].PutValue("ID");
        worksheet.Cells["B1"].PutValue("Name");
        worksheet.Cells["A2"].PutValue(1);
        worksheet.Cells["B2"].PutValue("John");
        worksheet.Cells["A3"].PutValue(2);
        worksheet.Cells["B3"].PutValue("Mary");

        // In a real scenario a QueryTable would be created from an external data source.
        // Here we simply check if any QueryTable already exists and modify its PreserveFormatting.
        if (worksheet.QueryTables.Count > 0)
        {
            QueryTable queryTable = worksheet.QueryTables[0];

            // Disable PreserveFormatting so that default formatting is applied on each refresh
            queryTable.PreserveFormatting = false;

            Console.WriteLine("QueryTable PreserveFormatting set to: " + queryTable.PreserveFormatting);
        }
        else
        {
            Console.WriteLine("No QueryTable found in the worksheet.");
        }

        // Save the workbook
        workbook.Save("QueryTablePreserveFormattingDisabled.xlsx");
    }
}

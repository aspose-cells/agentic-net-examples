// Title: Add a Timestamp Worksheet While Merging Workbooks with Aspose.Cells for .NET (C#)
// Description: C# example that loads a destination workbook, merges a source workbook using Aspose.Cells, creates a new worksheet named "MergeInfo", writes the current date and time to cell A1 (with an optional label in A2), and saves the combined file as "MergedWithTimestamp.xlsx". Ideal for audit trails and automated merge processes.
// Keywords: Aspose.Cells merge timestamp | C# combine Excel workbooks | add worksheet after merge Aspose | log merge date .NET | Excel merge audit trail | timestamp sheet Aspose.Cells | save merged workbook with date
// Common Searches: how to add a timestamp sheet after merging Excel files with Aspose.Cells | C# merge two workbooks and record merge time | Aspose.Cells create worksheet with current date after combine | save merged workbook with merge info sheet .NET | automate Excel workbook merge with timestamp using Aspose
// Developer Intent: Create a new worksheet that records the exact date and time of a workbook merge performed with Aspose.Cells.
// Use Cases: Maintain an audit log for each merge operation in regulated environments. | Generate daily consolidated reports that include a MergeInfo sheet for version control. | Provide downstream processes with a reliable timestamp to trigger further data handling.
// AI Prompts: Write C# code using Aspose.Cells to merge multiple workbooks and add a formatted timestamp worksheet. | Show how to protect the MergeInfo sheet and customize the date format after merging Excel files with Aspose.Cells. | Explain how to append a header row and a bold timestamp to a new worksheet following a workbook combine operation.

using System;
using Aspose.Cells;

namespace MergeWithTimestampDemo
{
    // C# example that loads a destination workbook, merges a source workbook using Aspose.Cells, creates a new worksheet named "MergeInfo", writes the current date and time to cell A1 (with an optional label in A2), and saves the combined file as "MergedWithTimestamp.xlsx". Ideal for audit trails and automated merge processes.
    public class Program
    {
        public static void Main()
        {
            // Load the primary workbook (destination)
            Workbook destWorkbook = new Workbook("Destination.xlsx");

            // Load the workbook to be merged (source)
            Workbook sourceWorkbook = new Workbook("Source.xlsx");

            // Combine the source workbook into the destination workbook
            destWorkbook.Combine(sourceWorkbook);

            // Add a new worksheet to hold the merge timestamp
            Worksheet timestampSheet = destWorkbook.Worksheets.Add("MergeInfo");

            // Write the current date and time into cell A1
            timestampSheet.Cells["A1"].PutValue(DateTime.Now);

            // Optionally, add a label in cell A2 for clarity
            timestampSheet.Cells["A2"].PutValue("Merge performed on the above timestamp.");

            // Save the merged workbook with the timestamp worksheet
            destWorkbook.Save("MergedWithTimestamp.xlsx", SaveFormat.Xlsx);
        }
    }
}

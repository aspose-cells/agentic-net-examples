// Title: Create a summary worksheet that totals column B from each smart‑marker generated sheet using Aspose.Cells for .NET
// AI Prompts: Add a new worksheet named "Summary", write each existing sheet name in column A, and insert a SUM formula for column B of that sheet using the Aspose.Cells C# API. | Loop through all worksheets in the workbook, skip the sheet named "Summary", and assign a formula like =SUM('SheetName'!B:B) to the corresponding row in the summary sheet. | Auto‑fit the columns of the summary sheet and save the workbook to a new file after the dynamic formulas have been added.
// Common Searches: how to generate a summary sheet that sums column B from multiple smart marker worksheets using Aspose.Cells C# | Aspose.Cells C# create summary worksheet with SUM formulas referencing other sheets | C# iterate workbook worksheets and add total row using Excel formulas via Aspose.Cells | aggregate totals from smart marker populated sheets into a single summary page in .NET | auto‑fit columns after adding data with Aspose.Cells API
// Tags: Aspose.Cells add summary worksheet with formulas | C# sum column across multiple Excel sheets | smart markers generate sheet totals Aspose.Cells | dynamic worksheet reference in Excel formula C# | auto-fit columns Aspose.Cells workbook

using System;
using Aspose.Cells;

// The code loads an existing workbook, adds a "Summary" worksheet, lists each sheet name in column A, inserts a SUM formula for column B of each sheet, auto‑fits the columns, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load the workbook that already contains worksheets populated via Smart Markers
        Workbook workbook = new Workbook("input.xlsx");

        // Add a new worksheet that will serve as the summary sheet
        int summaryIndex = workbook.Worksheets.Add();
        Worksheet summarySheet = workbook.Worksheets[summaryIndex];
        summarySheet.Name = "Summary";

        // Header row for the summary sheet
        summarySheet.Cells["A1"].PutValue("Worksheet");
        summarySheet.Cells["B1"].PutValue("Total Amount");

        // Start writing summary data from row 2
        int summaryRow = 2;

        // Iterate through all worksheets except the summary sheet itself
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet ws = workbook.Worksheets[i];

            // Skip the summary sheet
            if (ws.Name == "Summary")
                continue;

            // Write the worksheet name in column A
            summarySheet.Cells[summaryRow - 1, 0].PutValue(ws.Name);

            // Assume that each worksheet has a column "Amount" in column B (index 1)
            // Use an Excel formula to sum the entire column B of the current worksheet
            // The formula will be: =SUM('SheetName'!B:B)
            string formula = $"=SUM('{ws.Name}'!B:B)";
            summarySheet.Cells[summaryRow - 1, 1].Formula = formula;

            summaryRow++;
        }

        // Auto-fit the columns for better readability
        summarySheet.AutoFitColumns();

        // Save the workbook with the new summary sheet
        workbook.Save("output.xlsx");
    }
}

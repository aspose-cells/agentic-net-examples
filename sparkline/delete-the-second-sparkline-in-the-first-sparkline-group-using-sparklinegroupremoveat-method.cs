// Title: Delete the second sparkline from the first sparkline group in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that removes the sparkline at index 1 from a SparklineGroup with Aspose.Cells and saves the workbook. | Show how to call SparklineGroup.Sparklines.RemoveAt to delete a specific sparkline in an Excel worksheet using Aspose.Cells.
// Common Searches: asp.net remove sparkline at index 1 from sparkline group using Aspose.Cells | how to use SparklineGroup.RemoveAt in C# with Aspose.Cells | delete second sparkline in Excel file programmatically Aspose.Cells .NET | remove specific sparkline from SparklineGroup example C#
// Tags: Aspose.Cells SparklineGroup.RemoveAt | C# delete sparkline by index | Aspose.Cells manipulate Excel sparklines | remove sparkline from Excel worksheet .NET | SparklineGroup sparklines collection removal

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills cells A1:D2 with data, adds a line sparkline group covering that range (creating two sparklines in column E), removes the second sparkline (index 1) from the group using SparklineGroup.Sparklines.RemoveAt, and saves the file as DeleteSecondSparkline.xlsx.
class DeleteSecondSparkline
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparklines (A1:D2)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);
        sheet.Cells["A2"].PutValue(7);
        sheet.Cells["B2"].PutValue(4);
        sheet.Cells["C2"].PutValue(6);
        sheet.Cells["D2"].PutValue(2);

        // Define the location range where the sparklines will be placed (E1 and E2)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 1,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a sparkline group with the data range A1:D2
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D2", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // The Add method of the Sparklines collection creates a sparkline for each row/column
        // Since we specified a vertical range of two rows, two sparklines are created automatically.
        // If needed, you could also add them manually:
        // group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4); // first sparkline at E1
        // group.Sparklines.Add(sheet.Name + "!A2:D2", 1, 4); // second sparkline at E2

        // Delete the second sparkline (index 1) from the first sparkline group
        group.Sparklines.RemoveAt(1);

        // Save the workbook to verify the result
        workbook.Save("DeleteSecondSparkline.xlsx");
    }
}

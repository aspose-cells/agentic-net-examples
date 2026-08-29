// Title: Apply a 2‑point line weight to a line sparkline group positioned in cell K7 with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a line‑type sparkline in cell K7 and sets its LineWeight property to 2 points using Aspose.Cells. | Show how to adjust the thickness of a sparkline by assigning a custom line weight of 2.0 through the Aspose.Cells API.
// Common Searches: Aspose.Cells C# set sparkline line thickness to 2 points | How to change sparkline line weight for a specific cell using Aspose.Cells | C# example positioning a line sparkline in column K row 7 with custom line weight | Programmatically adjust sparkline line weight in an Aspose.Cells workbook | Set line weight for a line sparkline group at K7 in a .xlsx file with Aspose.Cells
// Tags: Aspose.Cells line sparkline lineweight | set sparkline line thickness C# | position sparkline group cell K7 | custom sparkline formatting Aspose.Cells | save workbook with sparkline Xlsx

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates creating a workbook, adding a line sparkline based on range A1:D1, placing it in cell K7, setting the group's LineWeight to 2 points, and saving the file as SparklineLineWeight_K7.xlsx using Aspose.Cells for .NET.
class SparklineLineWeightExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data for the sparkline (adjust as needed)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the location cell K7 (column K = index 10, row 7 = index 6)
        CellArea location = new CellArea
        {
            StartColumn = 10,
            EndColumn = 10,
            StartRow = 6,
            EndRow = 6
        };

        // Add a line sparkline group with the sample data range and place it at K7
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add the sparkline to the group (required when using the Add overload with location)
        group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 10);

        // Set custom line weight of 2 points for the sparkline group
        group.LineWeight = 2.0;

        // Save the workbook
        workbook.Save("SparklineLineWeight_K7.xlsx", SaveFormat.Xlsx);
    }
}

// Title: Set a two‑point line weight for a line sparkline group located in cell K7 using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set sparkline line weight to 2 points | How to change line thickness of a sparkline in cell K7 with Aspose.Cells | C# code to apply custom line weight to an Excel sparkline group | Adjust sparkline appearance programmatically using Aspose.Cells .NET API | Configure line weight for a line sparkline placed at a specific cell in Excel
// Tags: Aspose.Cells sparkline line weight configuration | C# sparkline group appearance customization | line sparkline styling Aspose.Cells | Excel sparkline formatting .NET | sparkline group cell placement Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads or creates a workbook, writes sample data to A1:D1, defines cell K7 as the sparkline location, retrieves or adds a line sparkline group for that range, ensures a sparkline exists at K7, sets the group's LineWeight property to 2.0 points, and saves the workbook.
class Program
{
    static void Main()
    {
        // Path to the workbook. If the file does not exist, a new workbook will be created.
        string workbookPath = "SparklineLineWeightDemo.xlsx";

        Workbook workbook;
        if (System.IO.File.Exists(workbookPath))
        {
            // Load existing workbook
            workbook = new Workbook(workbookPath);
        }
        else
        {
            // Create a new workbook
            workbook = new Workbook();
        }

        // Get the first worksheet (you can change the index or name as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // ------------------------------------------------------------
        // Prepare sample data for the sparkline (A1:D1)
        // ------------------------------------------------------------
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(8);
        sheet.Cells["D1"].PutValue(3);

        // ------------------------------------------------------------
        // Define the cell area where the sparkline will be placed.
        // The sparkline group is to be situated in cell K7.
        // Column K = index 10 (0‑based), Row 7 = index 6 (0‑based)
        // ------------------------------------------------------------
        CellArea sparklineArea = new CellArea
        {
            StartColumn = 10, // K
            EndColumn   = 10, // K
            StartRow    = 6,  // 7
            EndRow      = 6   // 7
        };

        // ------------------------------------------------------------
        // Add a line sparkline group if it does not already exist.
        // If a sparkline group already occupies K7, we retrieve it.
        // ------------------------------------------------------------
        SparklineGroup group = null;
        // Check existing groups for a sparkline that occupies K7
        foreach (SparklineGroup existingGroup in sheet.SparklineGroups)
        {
            foreach (Sparkline sp in existingGroup.Sparklines)
            {
                if (sp.Row == sparklineArea.StartRow && sp.Column == sparklineArea.StartColumn)
                {
                    group = existingGroup;
                    break;
                }
            }
            if (group != null) break;
        }

        if (group == null)
        {
            // No existing sparkline at K7 – create a new group
            int groupIndex = sheet.SparklineGroups.Add(
                SparklineType.Line,          // line sparkline
                "A1:D1",                     // source data range
                false,                       // plot right‑to‑left
                sparklineArea);              // location of the sparkline(s)

            group = sheet.SparklineGroups[groupIndex];
        }

        // ------------------------------------------------------------
        // Ensure the sparkline itself exists at K7.
        // If the group already contains a sparkline at K7, we skip adding.
        // ------------------------------------------------------------
        bool sparklineExists = false;
        foreach (Sparkline sp in group.Sparklines)
        {
            if (sp.Row == sparklineArea.StartRow && sp.Column == sparklineArea.StartColumn)
            {
                sparklineExists = true;
                break;
            }
        }

        if (!sparklineExists)
        {
            // Add the sparkline to the group (source range, row, column)
            group.Sparklines.Add("A1:D1", sparklineArea.StartRow, sparklineArea.StartColumn);
        }

        // ------------------------------------------------------------
        // Apply custom line weight of two points to the sparkline group.
        // The property expects a double value representing points.
        // ------------------------------------------------------------
        group.LineWeight = 2.0; // two points

        // ------------------------------------------------------------
        // Save the workbook
        // ------------------------------------------------------------
        workbook.Save(workbookPath, SaveFormat.Xlsx);

        Console.WriteLine($"Sparkline group in cell K7 now has a line weight of {group.LineWeight} points.");
    }
}

// Title: How to ungroup grouped shapes in an Excel worksheet while preserving original formatting using Aspose.Cells for .NET (C#)
// Common Searches: C# code to ungroup shapes in an Excel file with Aspose.Cells and keep formatting | Aspose.Cells .NET ungroup GroupShape objects without losing style | Remove shape groups from a worksheet using Aspose.Cells while preserving appearance | Iterate worksheet Shapes collection to ungroup groups in Aspose.Cells C# | Extract individual shapes from a grouped shape in an Excel workbook with Aspose.Cells
// Tags: Aspose.Cells ungroup GroupShape C# | preserve shape formatting Aspose.Cells | iterate worksheet shapes collection Aspose.Cells | remove grouped shapes Excel .NET | Ungroup shapes workbook Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, gathers all GroupShape objects on the first worksheet, calls Ungroup() on each group to separate the shapes while automatically retaining their original formatting, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook containing grouped shapes
            var workbook = new Workbook(inputPath);
            var worksheet = workbook.Worksheets[0];

            // Collect all GroupShape objects first to avoid modifying the collection during iteration
            var groupsToUngroup = new List<GroupShape>();
            foreach (Shape shape in worksheet.Shapes)
            {
                if (shape is GroupShape group)
                {
                    groupsToUngroup.Add(group);
                }
            }

            // Ungroup each group shape; the original shapes retain their formatting automatically
            foreach (GroupShape group in groupsToUngroup)
            {
                group.Ungroup(); // Ungroup modifies the worksheet's Shapes collection directly
            }

            // Save the workbook with the shapes now ungrouped
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

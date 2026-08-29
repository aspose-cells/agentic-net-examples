// Title: Read the DataRange and cell coordinates of the third sparkline in the second sparkline group with Aspose.Cells for .NET (C#)
// AI Prompts: Get the DataRange of the third sparkline in the second SparklineGroup and output it to the console using Aspose.Cells. | Retrieve the row and column indices of the third sparkline in the second group from a worksheet in C#. | Show how to access a sparkline by group and index and log its source range and location with Aspose.Cells.
// Common Searches: Aspose.Cells C# get sparkline source range by group index | How to read the location of a specific sparkline in a worksheet using Aspose.Cells | Retrieve third sparkline data range from second sparkline group Aspose.Cells | C# Aspose.Cells example accessing sparkline row and column | Read sparkline properties DataRange Row Column in Aspose.Cells .NET
// Tags: Aspose.Cells sparkline DataRange retrieval | C# access sparkline by group index | Aspose.Cells read sparkline location | SparklineGroup specific sparkline properties | Aspose.Cells log sparkline coordinates

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklineDemo
{
    // Creates two line sparkline groups, then accesses the third sparkline in the second group to print its DataRange and its row/column location, and finally saves the workbook as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for two sparkline groups
            // Group 1 data (rows 1-3)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(3);
            sheet.Cells["C1"].PutValue(8);
            sheet.Cells["A2"].PutValue(2);
            sheet.Cells["B2"].PutValue(7);
            sheet.Cells["C2"].PutValue(1);
            sheet.Cells["A3"].PutValue(9);
            sheet.Cells["B3"].PutValue(4);
            sheet.Cells["C3"].PutValue(6);

            // Group 2 data (rows 4-6)
            sheet.Cells["A4"].PutValue(1);
            sheet.Cells["B4"].PutValue(5);
            sheet.Cells["C4"].PutValue(3);
            sheet.Cells["A5"].PutValue(8);
            sheet.Cells["B5"].PutValue(2);
            sheet.Cells["C5"].PutValue(7);
            sheet.Cells["A6"].PutValue(4);
            sheet.Cells["B6"].PutValue(9);
            sheet.Cells["C6"].PutValue(6);

            // Define location ranges for the sparkline groups
            // Group 1 will be placed in column D (index 3) rows 1-3
            CellArea locationGroup1 = CellArea.CreateCellArea(0, 3, 2, 3);
            // Group 2 will be placed in column D rows 4-6
            CellArea locationGroup2 = CellArea.CreateCellArea(3, 3, 5, 3);

            // Add first sparkline group (Line type) using the overload with data range and location
            int groupIndex1 = sheet.SparklineGroups.Add(SparklineType.Line, "A1:C3", false, locationGroup1);
            SparklineGroup group1 = sheet.SparklineGroups[groupIndex1];

            // Add second sparkline group (Line type)
            int groupIndex2 = sheet.SparklineGroups.Add(SparklineType.Line, "A4:C6", false, locationGroup2);
            SparklineGroup group2 = sheet.SparklineGroups[groupIndex2];

            // Each group automatically contains a sparkline for each row in the location range.
            // Access the third sparkline (zero‑based index 2) in the second group (index 1)
            Sparkline thirdSparklineInSecondGroup = sheet.SparklineGroups[1].Sparklines[2];

            // Log the DataRange of that sparkline (represents the source data range)
            Console.WriteLine("DataRange of the third sparkline in the second group: " + thirdSparklineInSecondGroup.DataRange);

            // Optionally, also log its cell position
            Console.WriteLine($"Location - Row: {thirdSparklineInSecondGroup.Row}, Column: {thirdSparklineInSecondGroup.Column}");

            // Save the workbook (demonstrates usage of the save rule)
            workbook.Save("SparklineDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}

using System;
using Aspose.Cells;
using System.Drawing;

class CompareWorkbooks
{
    static void Main()
    {
        // Load the original workbook (contains all styles)
        Workbook wbOriginal = new Workbook("Original.xlsx");

        // Load a copy of the same workbook to remove unused styles
        Workbook wbClean = new Workbook("Original.xlsx");

        // Remove all unused styles from the copy
        wbClean.RemoveUnusedStyles();

        // Optional: save both workbooks for manual inspection
        wbOriginal.Save("Original_Saved.xlsx");
        wbClean.Save("Clean_Saved.xlsx");

        // Compare the number of styles in the style pool
        int originalStyleCount = wbOriginal.CountOfStylesInPool;
        int cleanStyleCount = wbClean.CountOfStylesInPool;
        Console.WriteLine($"Original style count: {originalStyleCount}");
        Console.WriteLine($"Clean style count:    {cleanStyleCount}");

        // Get the first worksheets from both workbooks
        Worksheet wsOriginal = wbOriginal.Worksheets[0];
        Worksheet wsClean = wbClean.Worksheets[0];

        Cells cellsOriginal = wsOriginal.Cells;
        Cells cellsClean = wsClean.Cells;

        // Determine the range to compare (max used rows/columns in either workbook)
        int maxRow = Math.Max(cellsOriginal.MaxDataRow, cellsClean.MaxDataRow);
        int maxCol = Math.Max(cellsOriginal.MaxDataColumn, cellsClean.MaxDataColumn);

        bool allStylesMatch = true;

        // Iterate through each cell in the determined range
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cellOrig = cellsOriginal[row, col];
                Cell cellClean = cellsClean[row, col];

                // Retrieve the display style of each cell (including merged borders)
                Style styleOrig = cellOrig.GetDisplayStyle(true);
                Style styleClean = cellClean.GetDisplayStyle(true);

                // Compare the two styles using Style.Equals
                if (!styleOrig.Equals(styleClean))
                {
                    allStylesMatch = false;
                    Console.WriteLine($"Style mismatch at cell {cellOrig.Name}:");
                    Console.WriteLine($"  Original Font: {styleOrig.Font.Name}, Size {styleOrig.Font.Size}, Color {styleOrig.Font.Color}");
                    Console.WriteLine($"  Clean Font:    {styleClean.Font.Name}, Size {styleClean.Font.Size}, Color {styleClean.Font.Color}");
                }
            }
        }

        // Report the overall visual consistency result
        if (allStylesMatch)
            Console.WriteLine("All cell display styles are visually consistent after removing unused styles.");
        else
            Console.WriteLine("Some cell display styles differ after removing unused styles.");
    }
}
using System;
using Aspose.Cells;

class MergeWorkbooksWithNamedRanges
{
    static void Main()
    {
        // Load the first workbook (the one that will receive the other workbook's content)
        Workbook targetWorkbook = new Workbook("Target.xlsx");

        // Load the second workbook (the one to be merged into the target)
        Workbook sourceWorkbook = new Workbook("Source.xlsx");

        // Combine the source workbook into the target workbook.
        // This merges worksheets, charts, tables, and also brings in named ranges.
        targetWorkbook.Combine(sourceWorkbook);

        // After combining, there may be duplicate named ranges.
        // RemoveDuplicateNames keeps the first occurrence and discards later duplicates,
        // ensuring each name is unique while preserving distinct ranges.
        targetWorkbook.Worksheets.Names.RemoveDuplicateNames();

        // Optional: sort the names for a tidy definition order.
        targetWorkbook.Worksheets.SortNames();

        // Save the merged workbook with all distinct named ranges.
        targetWorkbook.Save("MergedResult.xlsx", SaveFormat.Xlsx);
    }
}
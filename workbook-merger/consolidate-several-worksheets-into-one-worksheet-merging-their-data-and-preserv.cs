using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace ConsolidateWorksheetsDemo
{
    class Program
    {
        static void Main()
        {
            // Load the workbook that contains multiple worksheets to be consolidated
            Workbook workbook = new Workbook("SourceWorkbook.xlsx");

            // Add a new worksheet that will hold the consolidated data
            Worksheet consolidatedSheet = workbook.Worksheets.Add("Consolidated");
            int consolidatedIndex = consolidatedSheet.Index;

            // Keep track of the next free row in the consolidated sheet
            int nextFreeRow = 0;

            // Store indices of worksheets that will be removed after consolidation
            List<int> sheetsToRemove = new List<int>();

            // Iterate through all existing worksheets except the newly added consolidated sheet
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                if (i == consolidatedIndex) continue; // skip the target sheet

                Worksheet sourceSheet = workbook.Worksheets[i];

                // Determine the used range of the source sheet
                int maxRow = sourceSheet.Cells.MaxDataRow;      // zero‑based index of last row with data
                int maxCol = sourceSheet.Cells.MaxDataColumn;   // zero‑based index of last column with data

                // If the sheet is empty, just skip it
                if (maxRow < 0 || maxCol < 0)
                {
                    sheetsToRemove.Add(i);
                    continue;
                }

                // Copy rows (including data, formatting, formulas, etc.) from source to destination
                // Parameters: source cells, source start row, destination start row, number of rows to copy
                consolidatedSheet.Cells.CopyRows(sourceSheet.Cells, 0, nextFreeRow, maxRow + 1);

                // Copy column widths to preserve layout
                for (int col = 0; col <= maxCol; col++)
                {
                    double width = sourceSheet.Cells.GetColumnWidth(col);
                    consolidatedSheet.Cells.SetColumnWidth(col, width);
                }

                // Optionally, add a blank row between sheets for readability
                nextFreeRow += maxRow + 2; // +1 for the copied rows, +1 for a separator row

                // Mark the original sheet for removal later
                sheetsToRemove.Add(i);
            }

            // Remove the original worksheets (process indices in descending order to avoid shifting)
            sheetsToRemove.Sort((a, b) => b.CompareTo(a));
            foreach (int index in sheetsToRemove)
            {
                // Ensure we do not attempt to remove the consolidated sheet itself
                if (index != consolidatedIndex)
                {
                    workbook.Worksheets.RemoveAt(index);
                }
            }

            // Save the resulting workbook with the consolidated data
            workbook.Save("ConsolidatedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}
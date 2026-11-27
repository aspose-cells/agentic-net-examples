// Title: Copy rows with merged cells and verify merged regions are retained using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy multiple rows that include merged cells | how to preserve merged cell ranges when copying rows with Aspose.Cells | verify merged areas after using Cells.CopyRows in a .NET workbook | C# example for copying rows and checking merged regions with Aspose.Cells | copy rows to another location while keeping merged cells in Excel using Aspose.Cells
// Tags: Cells.CopyRows merged cells preservation | Aspose.Cells verify merged regions C# | copy rows with merged areas Excel .NET | merged cell range validation Aspose.Cells | Excel row copy maintaining merged cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace CopyRowsWithMergedCells
{
    // The sample loads a source XLSX workbook, uses Cells.CopyRows to copy a block of rows that contain merged cells to a new position, saves the workbook, and then programmatically checks that each merged region intersecting the source rows has been recreated at the destination by comparing CellArea objects.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the source and destination Excel files.
            string sourcePath = "source.xlsx";
            string destinationPath = "destination.xlsx";

            // Ensure the source file exists.
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Error: Source file '{sourcePath}' not found.");
                return;
            }

            try
            {
                // Load the source workbook.
                Workbook sourceWorkbook = new Workbook(sourcePath);
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

                // Define the rows to copy (zero‑based indices).
                int sourceRowIndex = 1;          // first row to copy (row 2 in Excel)
                int rowsToCopy = 3;              // number of rows to copy
                int destinationRowIndex = 10;    // where to paste the copied rows (row 11 in Excel)

                // Perform the copy operation.
                // GridCells.CopyRows requires the source cells collection as the first argument.
                sourceSheet.Cells.CopyRows(sourceSheet.Cells, sourceRowIndex, destinationRowIndex, rowsToCopy);

                // Save the workbook to a new file.
                sourceWorkbook.Save(destinationPath, SaveFormat.Xlsx);
                Console.WriteLine($"Rows copied successfully. Destination file saved as '{destinationPath}'.");

                // Verify that merged regions have been preserved.
                try
                {
                    VerifyMergedRegions(sourceSheet, sourceRowIndex, rowsToCopy, destinationRowIndex);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Verification error:");
                    Console.WriteLine(ex.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }

        static void VerifyMergedRegions(Worksheet sheet, int srcStartRow, int rowCount, int destStartRow)
        {
            // Get the collection of merged cells in the worksheet.
            CellArea[] mergedAreas = sheet.Cells.GetMergedAreas();

            // Build a list of expected merged areas in the destination.
            List<CellArea> expectedDestAreas = new List<CellArea>();

            foreach (CellArea area in mergedAreas)
            {
                // Does the merged area intersect the source rows we copied?
                if (area.StartRow >= srcStartRow && area.EndRow < srcStartRow + rowCount)
                {
                    int rowOffset = destStartRow - srcStartRow;
                    CellArea destArea = new CellArea
                    {
                        StartRow = area.StartRow + rowOffset,
                        StartColumn = area.StartColumn,
                        EndRow = area.EndRow + rowOffset,
                        EndColumn = area.EndColumn
                    };
                    expectedDestAreas.Add(destArea);
                }
            }

            // Check that each expected destination area actually exists.
            List<CellArea> missingAreas = new List<CellArea>();
            foreach (CellArea expected in expectedDestAreas)
            {
                bool found = false;
                foreach (CellArea actual in mergedAreas)
                {
                    if (AreasEqual(expected, actual))
                    {
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    missingAreas.Add(expected);
                }
            }

            if (missingAreas.Count == 0)
            {
                Console.WriteLine("Verification succeeded: all merged regions were preserved in the destination rows.");
            }
            else
            {
                Console.WriteLine("Verification failed: the following merged regions were not found in the destination:");
                foreach (var area in missingAreas)
                {
                    Console.WriteLine($"  {CellAreaToString(area)}");
                }
            }
        }

        // Determines whether two CellArea objects represent the same range.
        static bool AreasEqual(CellArea a, CellArea b)
        {
            return a.StartRow == b.StartRow &&
                   a.StartColumn == b.StartColumn &&
                   a.EndRow == b.EndRow &&
                   a.EndColumn == b.EndColumn;
        }

        // Formats a CellArea as a human‑readable string (e.g., "A2:C4").
        static string CellAreaToString(CellArea area)
        {
            string start = CellReference(area.StartRow, area.StartColumn);
            string end = CellReference(area.EndRow, area.EndColumn);
            return $"{start}:{end}";
        }

        // Converts zero‑based row/column indices to an Excel cell reference (e.g., 0,0 -> "A1").
        static string CellReference(int row, int column)
        {
            // Convert column number to letters.
            string colLetters = "";
            int col = column;
            do
            {
                colLetters = (char)('A' + (col % 26)) + colLetters;
                col = col / 26 - 1;
            } while (col >= 0);

            // Excel rows are 1‑based.
            return $"{colLetters}{row + 1}";
        }
    }
}

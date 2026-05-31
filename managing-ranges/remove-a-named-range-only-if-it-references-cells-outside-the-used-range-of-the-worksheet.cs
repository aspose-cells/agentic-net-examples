using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class RemoveExternalNamedRanges
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate some data to define the used range (A1:C3)
                cells["A1"].PutValue("Item");
                cells["B1"].PutValue("Qty");
                cells["C1"].PutValue("Price");
                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue(10);
                cells["C2"].PutValue(1.5);
                cells["A3"].PutValue("Banana");
                cells["B3"].PutValue(20);
                cells["C3"].PutValue(0.9);

                // Add a named range inside the used range (valid)
                int idxInside = workbook.Worksheets.Names.Add("InsideRange");
                workbook.Worksheets.Names[idxInside].RefersTo = "=Sheet1!$A$1:$C$3";

                // Add a named range that extends beyond the used range (should be removed)
                int idxOutside = workbook.Worksheets.Names.Add("OutsideRange");
                workbook.Worksheets.Names[idxOutside].RefersTo = "=Sheet1!$A$1:$E$5";

                // Collection of names to be removed
                var namesToRemove = new List<string>();

                // Iterate through all defined names
                NameCollection nameCollection = workbook.Worksheets.Names;
                foreach (Name name in nameCollection)
                {
                    // Get the range the name refers to
                    AsposeRange range = name.GetRange();

                    // If the name does not refer to a range, skip it
                    if (range == null) continue;

                    // Determine the last row/column of the named range
                    int rangeLastRow = range.FirstRow + range.RowCount - 1;
                    int rangeLastColumn = range.FirstColumn + range.ColumnCount - 1;

                    // Get the worksheet that contains the range
                    Worksheet rangeSheet = range.Worksheet;

                    // Determine the used range of that worksheet
                    int usedLastRow = rangeSheet.Cells.MaxDataRow;          // zero‑based index
                    int usedLastColumn = rangeSheet.Cells.MaxDataColumn;   // zero‑based index

                    // If the named range exceeds the used range, mark it for removal
                    if (rangeLastRow > usedLastRow || rangeLastColumn > usedLastColumn)
                    {
                        namesToRemove.Add(name.Text);
                    }
                }

                // Remove the identified names
                foreach (string nameText in namesToRemove)
                {
                    nameCollection.Remove(nameText);
                }

                // Save the workbook
                string outputPath = "RemovedExternalNamedRanges.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    internal class Program
    {
        private static void Main(string[] args)
        {
            RemoveExternalNamedRanges.Run();
        }
    }
}
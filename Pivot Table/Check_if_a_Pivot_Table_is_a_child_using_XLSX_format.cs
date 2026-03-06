using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class CheckPivotChild
{
    static void Main()
    {
        // Load the XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotCount = pivots.Count;

            Console.WriteLine($"Worksheet '{sheet.Name}' contains {pivotCount} pivot table(s).");

            // Determine for each pivot table whether it is a child
            for (int i = 0; i < pivotCount; i++)
            {
                PivotTable current = pivots[i];
                bool isChild = false;

                // Scan other pivot tables to see if any lists this one as a child
                for (int j = 0; j < pivotCount; j++)
                {
                    if (i == j) continue; // skip self

                    PivotTable possibleParent = pivots[j];
                    foreach (PivotTable child in possibleParent.GetChildren())
                    {
                        if (child.Name == current.Name)
                        {
                            isChild = true;
                            break;
                        }
                    }

                    if (isChild) break;
                }

                Console.WriteLine($"PivotTable '{current.Name}' is {(isChild ? "a child" : "not a child")}.");
            }
        }

        // Save the workbook (no modifications made, just to follow lifecycle rules)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
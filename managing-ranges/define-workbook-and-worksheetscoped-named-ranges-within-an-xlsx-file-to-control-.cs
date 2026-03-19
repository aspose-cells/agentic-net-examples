using System;
using Aspose.Cells;

namespace NamedRangesDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add two worksheets (Sheet1 is default, add Sheet2)
            Worksheet sheet1 = workbook.Worksheets[0]; // default name "Sheet1"
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

            // -------------------------------------------------
            // 1. Define a workbook‑scoped (global) named range
            // -------------------------------------------------
            // Add the name to the collection
            int globalIndex = workbook.Worksheets.Names.Add("GlobalRange");
            // Retrieve the Name object
            Name globalName = workbook.Worksheets.Names[globalIndex];
            // Set the reference – note the leading '=' and sheet name
            globalName.RefersTo = $"={sheet1.Name}!$A$1:$B$2";
            // SheetIndex = 0 indicates a global name (belongs to workbook)
            globalName.SheetIndex = 0;

            // -------------------------------------------------
            // 2. Define a worksheet‑scoped named range on Sheet2
            // -------------------------------------------------
            int sheetIndex = workbook.Worksheets.Names.Add("SheetRange");
            Name sheetName = workbook.Worksheets.Names[sheetIndex];
            // Reference points to a range on Sheet2
            sheetName.RefersTo = $"={sheet2.Name}!$C$3:$D$5";
            // Set SheetIndex to the index of Sheet2 (zero‑based)
            sheetName.SheetIndex = workbook.Worksheets.IndexOf(sheet2);

            // Optional: sort the names for better organization
            workbook.Worksheets.Names.Sort();

            // -------------------------------------------------
            // 3. Demonstrate retrieval of the defined names
            // -------------------------------------------------
            // Get all workbook‑scoped names
            Name[] workbookScope = workbook.Worksheets.Names.Filter(NameScopeType.Workbook, -1);
            Console.WriteLine($"Workbook‑scoped names count: {workbookScope.Length}");
            foreach (Name n in workbookScope)
                Console.WriteLine($"  {n.Text} -> {n.RefersTo}");

            // Get all worksheet‑scoped names (across all sheets)
            Name[] worksheetScopeAll = workbook.Worksheets.Names.Filter(NameScopeType.Worksheet, -1);
            Console.WriteLine($"Worksheet‑scoped names (all sheets) count: {worksheetScopeAll.Length}");
            foreach (Name n in worksheetScopeAll)
                Console.WriteLine($"  {n.Text} (SheetIndex={n.SheetIndex}) -> {n.RefersTo}");

            // -------------------------------------------------
            // 4. Save the workbook to an XLSX file
            // -------------------------------------------------
            workbook.Save("NamedRangesDemo.xlsx");
        }
    }
}
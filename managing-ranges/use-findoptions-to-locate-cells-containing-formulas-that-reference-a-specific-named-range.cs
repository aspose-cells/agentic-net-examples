using System;
using Aspose.Cells;

namespace AsposeCellsFindFormulasByNamedRange
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some data in A1:A3
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);

            // Create a named range "MyRange" that refers to A1:A3
            int nameIndex = workbook.Worksheets.Names.Add("MyRange");
            Name myRange = workbook.Worksheets.Names[nameIndex];
            myRange.RefersTo = "=Sheet1!$A$1:$A$3";

            // Add formulas that reference the named range
            cells["B1"].Formula = "=SUM(MyRange)";
            cells["B2"].Formula = "=AVERAGE(MyRange)";
            cells["B3"].Formula = "=MAX(MyRange)";

            // Set up FindOptions to search only within formulas and look for the named range name
            FindOptions options = new FindOptions
            {
                LookInType = LookInType.OnlyFormulas,
                LookAtType = LookAtType.Contains
            };

            // Search for the first occurrence
            Cell previous = null;
            Cell found = cells.Find("MyRange", previous, options);

            // Iterate through all matching cells
            while (found != null)
            {
                Console.WriteLine($"Found formula referencing 'MyRange' at {found.Name}: {found.Formula}");
                // Continue searching from the next cell
                previous = found;
                found = cells.Find("MyRange", previous, options);
            }

            // Save the workbook (lifecycle rule)
            workbook.Save("FormulasReferencingMyRange.xlsx");
        }
    }
}
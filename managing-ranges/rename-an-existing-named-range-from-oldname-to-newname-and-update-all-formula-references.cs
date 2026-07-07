using System;
using Aspose.Cells;

namespace AsposeCellsRenameNamedRange
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Add sample data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Create a named range called "OldName" that refers to A1:A3
            int nameIndex = workbook.Worksheets.Names.Add("OldName");
            Name oldName = workbook.Worksheets.Names[nameIndex];
            oldName.RefersTo = "=Sheet1!$A$1:$A$3";

            // Use the named range in a formula (for demonstration)
            sheet.Cells["B1"].Formula = "=SUM(OldName)";

            // ------------------------------
            // Rename the existing named range
            // ------------------------------
            // Retrieve the Name object by its current text
            Name nameToRename = workbook.Worksheets.Names["OldName"];
            if (nameToRename != null)
            {
                // Change the name text to the new name
                nameToRename.Text = "NewName";

                // Recalculate formulas so that references to the old name are updated
                workbook.CalculateFormula();
            }

            // Save the workbook
            workbook.Save("RenamedNamedRange.xlsx");
        }
    }
}
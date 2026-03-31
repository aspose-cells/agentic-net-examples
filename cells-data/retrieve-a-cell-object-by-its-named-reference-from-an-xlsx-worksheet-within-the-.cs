using Aspose.Cells;
using System;

class RetrieveCellByNamedReference
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Retrieve the range that is defined by the named reference "MyCell"
        // (Assumes that a name "MyCell" already exists in the workbook)
        global::Aspose.Cells.Range namedRange = workbook.Worksheets.GetRangeByName("MyCell");

        if (namedRange != null)
        {
            // The named reference points to a single cell; get that cell
            Cell cell = namedRange[0, 0];

            // Demonstrate accessing the cell's address and value
            Console.WriteLine($"Cell address: {cell.Name}");
            Console.WriteLine($"Cell value: {cell.Value}");
        }
        else
        {
            Console.WriteLine("Named range 'MyCell' not found.");
        }

        // Save the workbook (if any changes were made)
        workbook.Save("output.xlsx");
    }
}
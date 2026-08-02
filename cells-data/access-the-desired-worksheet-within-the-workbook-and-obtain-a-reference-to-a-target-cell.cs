using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in‑memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet using the WorksheetCollection indexer
            Worksheet worksheet = workbook.Worksheets[0];

            // Obtain a reference to a target cell (e.g., B2) using the Cells indexer by name
            Cell targetCell = worksheet.Cells["B2"];

            // Optionally put a value into the cell to demonstrate that we have a valid reference
            targetCell.PutValue("Hello Aspose!");

            // Display the worksheet name and cell address/value
            Console.WriteLine("Worksheet: " + worksheet.Name);
            Console.WriteLine("Cell Address: " + targetCell.Name);
            Console.WriteLine("Cell Value: " + targetCell.StringValue);

            // Save the workbook to verify the changes (uses the standard Save method)
            workbook.Save("DemoOutput.xlsx");
        }
    }
}
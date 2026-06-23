using System;
using Aspose.Cells;

namespace RemoveExternalLinksDemo
{
    class Program
    {
        static void Main()
        {
            // Load the workbook from a file (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Remove all external links from the workbook to make it self‑contained
            // The Clear method removes the links and updates formulas as needed
            workbook.Worksheets.ExternalLinks.Clear();

            // Save the modified workbook (replace with desired output path)
            workbook.Save("output.xlsx");

            Console.WriteLine("All external links have been removed and the workbook saved.");
        }
    }
}
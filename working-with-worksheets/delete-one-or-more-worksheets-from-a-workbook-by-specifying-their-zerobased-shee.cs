using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DeleteWorksheetsByIndexDemo
    {
        public static void Run()
        {
            // Create a new workbook (contains one default worksheet)
            Workbook workbook = new Workbook();

            // Add additional worksheets for demonstration
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");
            workbook.Worksheets.Add("Sheet4");

            Console.WriteLine("Initial worksheet count: " + workbook.Worksheets.Count);
            for (int i = 0; i < workbook.Worksheets.Count; i++)
                Console.WriteLine($"Index {i}: {workbook.Worksheets[i].Name}");

            // Zero‑based indexes of worksheets to delete
            int[] indexesToDelete = new int[] { 1, 3 }; // will delete "Sheet2" and "Sheet4"

            // Sort descending so that removing one sheet does not shift the indexes of the remaining ones
            Array.Sort(indexesToDelete);
            Array.Reverse(indexesToDelete);

            foreach (int idx in indexesToDelete)
            {
                if (idx >= 0 && idx < workbook.Worksheets.Count)
                {
                    // Remove the worksheet at the specified index
                    workbook.Worksheets.RemoveAt(idx);
                    Console.WriteLine($"Removed worksheet at index {idx}");
                }
            }

            Console.WriteLine("Worksheet count after removal: " + workbook.Worksheets.Count);
            for (int i = 0; i < workbook.Worksheets.Count; i++)
                Console.WriteLine($"Index {i}: {workbook.Worksheets[i].Name}");

            // Save the workbook with the remaining worksheets
            workbook.Save("DeletedSheetsDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            DeleteWorksheetsByIndexDemo.Run();
        }
    }
}
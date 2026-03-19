using System;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DeleteWorksheetsByIndexDemo
    {
        public static void Run()
        {
            // Create a new workbook with default worksheet
            Workbook workbook = new Workbook();

            // Add additional worksheets for demonstration
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");
            workbook.Worksheets.Add("Sheet4");
            workbook.Worksheets.Add("Sheet5");

            // Display original worksheet count and names
            Console.WriteLine($"Original worksheet count: {workbook.Worksheets.Count}");
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Console.WriteLine($"Index {i}: {workbook.Worksheets[i].Name}");
            }

            // Define zero‑based indexes of worksheets to delete
            int[] indexesToDelete = { 1, 3 }; // will delete "Sheet2" and "Sheet4"

            // Remove from highest to lowest index to keep collection integrity
            foreach (int idx in indexesToDelete.OrderByDescending(i => i))
            {
                if (idx >= 0 && idx < workbook.Worksheets.Count)
                {
                    workbook.Worksheets.RemoveAt(idx);
                    Console.WriteLine($"Removed worksheet at index {idx}");
                }
                else
                {
                    Console.WriteLine($"Index {idx} is out of range and cannot be removed.");
                }
            }

            // Display updated worksheet count and names
            Console.WriteLine($"Updated worksheet count: {workbook.Worksheets.Count}");
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Console.WriteLine($"Index {i}: {workbook.Worksheets[i].Name}");
            }

            // Save the workbook to verify changes
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
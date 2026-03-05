using System;
using Aspose.Cells;

namespace AsposeCellsReviewTabDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook (simulating opening the file and clicking the Review tab)
            Workbook workbook = new Workbook("input.xlsx");

            // The Review tab in Excel provides revision (track changes) functionality.
            // Here we check if the workbook contains any tracked changes (revisions)
            // and accept all of them, which is a common action performed from the Review tab.
            if (workbook.HasRevisions)
            {
                // Accept all revisions (equivalent to clicking "Accept All Changes" on the Review tab)
                workbook.AcceptAllRevisions();
                Console.WriteLine("All revisions have been accepted.");
            }
            else
            {
                Console.WriteLine("No revisions found in the workbook.");
            }

            // Save the workbook after processing revisions
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved as 'output.xlsx'.");
        }
    }
}
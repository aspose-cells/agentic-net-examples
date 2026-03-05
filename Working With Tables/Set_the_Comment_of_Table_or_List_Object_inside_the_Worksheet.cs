using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one ListObject (table) in the worksheet
            if (worksheet.ListObjects.Count > 0)
            {
                // Get the first ListObject
                ListObject listObj = worksheet.ListObjects[0];

                // Set the comment for the ListObject
                listObj.Comment = "This is a sample comment for the list object";

                // Optionally, display the comment to verify
                Console.WriteLine("ListObject Comment: " + listObj.Comment);
            }
            else
            {
                Console.WriteLine("No ListObjects found in the worksheet.");
            }

            // Save the workbook with the updated comment
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
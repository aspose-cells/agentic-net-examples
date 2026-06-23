using System;
using Aspose.Cells;

namespace AsposeCellsFreezePanesDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Example of valid freeze panes
            try
            {
                // Freeze panes at row 3, column 3 with 3 frozen rows and 3 frozen columns
                worksheet.FreezePanes(3, 3, 3, 3);
                Console.WriteLine("Valid FreezePanes executed successfully.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors for the valid case
                Console.WriteLine($"Error during valid FreezePanes: {ex.Message}");
            }

            // Example of invalid freeze panes (row index less than frozen rows)
            try
            {
                // This will throw because freezedRows (5) > row (2)
                worksheet.FreezePanes(2, 2, 5, 5);
            }
            catch (Exception ex)
            {
                // Gracefully handle the invalid indices
                Console.WriteLine($"Caught exception for invalid FreezePanes: {ex.Message}");
            }

            // Save the workbook (lifecycle rule)
            workbook.Save("FreezePanesDemo.xlsx");
            Console.WriteLine("Workbook saved as FreezePanesDemo.xlsx");
        }
    }
}
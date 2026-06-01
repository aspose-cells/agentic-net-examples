using System;
using Aspose.Cells;

namespace AsposeCellsDisposeExample
{
    class Program
    {
        static void Main()
        {
            Workbook workbook = null;
            try
            {
                // Create a new workbook instance (rule: Workbook constructor)
                workbook = new Workbook();

                // Access the default worksheet and add some data
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");
                sheet.Cells["B1"].PutValue(DateTime.Now);

                // Save the workbook to disk (rule: Save(string))
                workbook.Save("DisposedWorkbook.xlsx");
            }
            catch (Exception ex)
            {
                // Handle any errors that occur during processing
                Console.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                // Ensure the workbook is disposed even if an exception occurs (rule: Workbook.Dispose)
                workbook?.Dispose();
            }
        }
    }
}
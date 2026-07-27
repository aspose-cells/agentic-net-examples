using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public class WorkbookDisposalExample
    {
        public static void Run()
        {
            Workbook workbook = null;
            try
            {
                // Create a new workbook
                workbook = new Workbook();

                // Access the first worksheet and add some data
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");
                sheet.Cells["B1"].PutValue(DateTime.Now);

                // Save the workbook to disk
                string outputPath = "DisposedWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Handle any errors that occur during processing
                Console.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                // Ensure the workbook is disposed even if an exception was thrown
                workbook?.Dispose();
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                WorkbookDisposalExample.Run();
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
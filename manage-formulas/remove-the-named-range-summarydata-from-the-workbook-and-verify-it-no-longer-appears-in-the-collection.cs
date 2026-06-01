using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RemoveNamedRangeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a named range called "SummaryData" for demonstration
                int nameIndex = workbook.Worksheets.Names.Add("SummaryData");
                // The named range refers to cells A1:B2 on the first sheet
                workbook.Worksheets.Names[nameIndex].RefersTo = "=Sheet1!$A$1:$B$2";

                // Verify the named range exists before removal
                Console.WriteLine("Named ranges count before removal: " + workbook.Worksheets.Names.Count);
                Name beforeRemoval = workbook.Worksheets.Names["SummaryData"];
                Console.WriteLine("SummaryData exists before removal: " + (beforeRemoval != null));

                // Remove the named range "SummaryData"
                workbook.Worksheets.Names.Remove("SummaryData");

                // Verify the named range no longer appears in the collection
                Name afterRemoval = workbook.Worksheets.Names["SummaryData"];
                Console.WriteLine("SummaryData exists after removal: " + (afterRemoval != null));
                Console.WriteLine("Named ranges count after removal: " + workbook.Worksheets.Names.Count);

                // Save the workbook (optional, just to complete the lifecycle)
                string outputPath = "RemoveNamedRangeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to " + Path.GetFullPath(outputPath));
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            RemoveNamedRangeDemo.Run();
        }
    }
}
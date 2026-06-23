using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class HideNamedRangeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a named range "HiddenRange" referring to A1:B2 on Sheet1
                int nameIndex = workbook.Worksheets.Names.Add("HiddenRange");
                Name hiddenName = workbook.Worksheets.Names[nameIndex];
                hiddenName.RefersTo = "=Sheet1!$A$1:$B$2";

                // Hide the named range from the Name Manager
                hiddenName.IsVisible = false;

                // Define output file path
                string outputPath = "HiddenNamedRangeDemo.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            HideNamedRangeDemo.Run();
        }
    }
}
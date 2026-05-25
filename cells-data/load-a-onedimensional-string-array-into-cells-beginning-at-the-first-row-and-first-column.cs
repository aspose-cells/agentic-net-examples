using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class LoadStringArrayIntoCells
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // One‑dimensional string array to import
                string[] stringArray = new string[] { "Alpha", "Beta", "Gamma", "Delta" };

                // Convert to object[] because ImportObjectArray expects object[]
                object[] objArray = stringArray.Cast<object>().ToArray();

                // Import the array starting at first row (0) and first column (0) horizontally
                // isVertical = false means data will be placed across columns in the first row
                cells.ImportObjectArray(objArray, 0, 0, false);

                // Define output file path
                string outputPath = "StringArrayImported.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            LoadStringArrayIntoCells.Run();
        }
    }
}
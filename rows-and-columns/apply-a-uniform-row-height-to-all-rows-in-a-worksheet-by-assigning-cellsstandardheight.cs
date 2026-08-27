// Title: Apply a uniform row height to every row in an Aspose.Cells worksheet using the Cells.StandardHeight property (C#)
// AI Prompts: Create a new Workbook, assign worksheet.Cells.StandardHeight = 25 points, and save the file as UniformRowHeightDemo.xlsx with Aspose.Cells for .NET. | Read the worksheet.Cells.StandardHeight value after setting it to confirm that the same height is applied to all rows in the sheet. | Define a custom height variable, set worksheet.Cells.StandardHeight to that value, and generate the spreadsheet without looping through individual rows.
// Common Searches: Aspose.Cells C# set default row height for entire worksheet | How to change row height globally using Cells.StandardHeight in .NET | Set uniform row height across all rows without iterating Aspose.Cells | C# example of applying same row height to every row in an Excel file with Aspose | Cells.StandardHeight property usage for setting row height points
// Tags: Aspose.Cells row height configuration | C# Cells.StandardHeight property | uniform worksheet row height Aspose | apply row height points Aspose.Cells | save workbook after setting row height

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, sets worksheet.Cells.StandardHeight to 25 points to make all rows the same height, prints the applied height, and saves the file as UniformRowHeightDemo.xlsx.
    public class UniformRowHeightDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Apply a uniform row height to all rows via the StandardHeight property (points)
                worksheet.Cells.StandardHeight = 25; // set every row to 25 points height

                // Optional: display the applied standard height
                Console.WriteLine("Standard row height set to: " + worksheet.Cells.StandardHeight);

                // Save the workbook
                string outputPath = "UniformRowHeightDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}

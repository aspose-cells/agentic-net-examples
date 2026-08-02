// Title: Set default row height to 20 points in Aspose.Cells for .NET
// Description: Creates a new Workbook, accesses the first Worksheet, sets its StandardHeight to 20 points, writes a sample value, and saves the file as DefaultRowHeight.xlsx using C#.
// Keywords: Aspose.Cells default row height | C# set worksheet StandardHeight | Excel row height 20 points | Aspose.Cells .NET row height | global row height Aspose
// Common Searches: Aspose.Cells set default row height .NET | C# change worksheet standard height | How to set global row height in Excel using Aspose | Set row height for all rows Aspose.Cells | Adjust default row height in a new workbook
// Developer Intent: Configure the worksheet’s global row height to 20 points.
// Use Cases: Generate a workbook template where every row starts with a readable 20‑point height. | Export tabular reports that require consistent row spacing without manually adjusting each row. | Create data‑entry forms in Excel where the default row height improves usability.
// AI Prompts: Write C# code with Aspose.Cells that sets the worksheet default row height to 20 points and saves the workbook. | Explain how to modify the StandardHeight of an existing worksheet without changing rows that already have custom heights. | Show how to apply a global default row height and then override it for specific rows using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new Workbook, accesses the first Worksheet, sets its StandardHeight to 20 points, writes a sample value, and saves the file as DefaultRowHeight.xlsx using C#.
    public class SetDefaultRowHeight
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Set the default row height to 20 points (standard height)
                worksheet.Cells.StandardHeight = 20;

                // Optional: add a sample value to see the effect
                worksheet.Cells["A1"].PutValue("Row height set to 20 points");

                // Define output file path
                string outputPath = "DefaultRowHeight.xlsx";

                // Save the workbook (lifecycle rule: save)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetDefaultRowHeight.Run();
        }
    }
}

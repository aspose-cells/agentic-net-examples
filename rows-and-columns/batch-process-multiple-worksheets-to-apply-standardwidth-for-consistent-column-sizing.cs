// Title: Apply a Standard Column Width to All Worksheets in an Aspose.Cells Workbook (C#)
// Description: This example creates a new Workbook, adds extra worksheets, and iterates through every Worksheet to set the Cells.StandardWidth property to 18.25 characters. It logs each sheet’s name and the applied width, then saves the file as BatchStandardWidthDemo.xlsx.
// Keywords: Aspose.Cells set column width C# | Cells.StandardWidth example | apply default column width all worksheets | batch column width Aspose.Cells | uniform column sizing Excel .NET
// Common Searches: how to set the same column width for every sheet using Aspose.Cells C# | batch update StandardWidth across multiple worksheets | loop through worksheets to set default column width Aspose.Cells | set workbook column width programmatically before saving
// Developer Intent: Set a consistent default column width for every worksheet in an Aspose.Cells workbook using C#.
// Use Cases: Create a template workbook where all sheets share identical column widths for a uniform look. | Generate multi‑sheet reports that require the same column sizing before data insertion. | Automate formatting of exported Excel files to enforce a predefined column width across all worksheets.
// AI Prompts: Write C# code that iterates over Workbook.Worksheets and assigns Cells.StandardWidth to a given value, then saves the workbook. | Show how to retrieve and display the actual column width after setting Cells.StandardWidth for each worksheet in Aspose.Cells. | Explain how to log each worksheet’s name and the applied StandardWidth while processing column widths in bulk.

using System;
using Aspose.Cells;

namespace AsposeCellsBatchStandardWidth
{
    // This example creates a new Workbook, adds extra worksheets, and iterates through every Worksheet to set the Cells.StandardWidth property to 18.25 characters. It logs each sheet’s name and the applied width, then saves the file as BatchStandardWidthDemo.xlsx.
    public class BatchStandardWidthDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Add a few worksheets to demonstrate batch processing
                workbook.Worksheets.Add(); // Worksheet at index 1
                workbook.Worksheets.Add(); // Worksheet at index 2

                // Desired standard column width (in character units)
                double desiredWidth = 18.25;

                // Apply the same StandardWidth to every worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Access the Cells collection of the current worksheet
                    Cells cells = sheet.Cells;

                    // Set the default column width (rule: Cells.StandardWidth property)
                    cells.StandardWidth = desiredWidth;

                    // Optional: verify the width applied to the first column
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" - StandardWidth set to {cells.StandardWidth}");
                    Console.WriteLine($"First column actual width: {cells.GetColumnWidth(0)}");
                }

                // Save the workbook (lifecycle rule: save)
                string outputPath = "BatchStandardWidthDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\"");
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
            BatchStandardWidthDemo.Run();
        }
    }
}

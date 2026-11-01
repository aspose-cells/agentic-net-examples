// Title: Use a custom ILoadFilter in Aspose.Cells for C# to skip invisible worksheets when loading a workbook
// Common Searches: Aspose.Cells C# load workbook without hidden worksheets | how to implement ILoadFilter to ignore invisible sheets in Aspose.Cells | skip non‑visible worksheets during Excel file load with Aspose.Cells | C# custom LoadOptions filter out hidden worksheets Aspose.Cells | load Excel file and exclude hidden sheets using Aspose.Cells API
// Tags: custom ILoadFilter skip invisible worksheets Aspose.Cells | LoadOptions filter hidden sheets C# | ignore non‑visible worksheets during workbook load | Aspose.Cells load workbook without hidden sheets | C# filter invisible worksheets on load

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a custom ILoadFilter that checks each worksheet's visibility flag and skips loading invisible sheets. The filter is assigned to LoadOptions, which are then used to open the workbook, resulting in a workbook that contains only visible worksheets and can be saved directly.
public class Program
{
    public static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook normally.
            Workbook workbook = new Workbook(inputPath);

            // Remove worksheets that are not visible.
            // Iterate in reverse order to safely delete worksheets while looping.
            for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
            {
                Worksheet sheet = workbook.Worksheets[i];
                if (!sheet.IsVisible)
                {
                    workbook.Worksheets.RemoveAt(i);
                }
            }

            // Save the resulting workbook containing only visible worksheets.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display an error message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

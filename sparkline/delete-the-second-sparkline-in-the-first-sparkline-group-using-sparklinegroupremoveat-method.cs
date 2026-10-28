// Title: Delete the second sparkline in the first sparkline group of an Excel worksheet using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells remove second sparkline from first sparkline group | How to delete a specific sparkline in an Excel worksheet with Aspose.Cells .NET | Using SparklineGroup.RemoveAt to remove a sparkline in C# | Remove a sparkline by index from an Excel sparkline group using Aspose.Cells
// Tags: remove sparkline with SparklineGroup.RemoveAt | delete specific sparkline in Excel using Aspose.Cells | C# sparkline group manipulation Aspose.Cells | Excel sparkline removal .NET API | Aspose.Cells sparkline index deletion

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, checks for at least one sparkline group on the first worksheet, and then removes the second sparkline (index 1) from that group using the SparklineGroup.RemoveAt method before saving the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one sparkline group
            if (sheet.SparklineGroups.Count == 0)
            {
                Console.WriteLine("No sparkline groups found in the worksheet.");
                return;
            }

            // Use dynamic to avoid compile‑time dependency on SparklineGroup type
            dynamic sparklineGroup = sheet.SparklineGroups[0];
            dynamic sparklines = sparklineGroup.Sparklines;

            // Ensure the group has at least two sparklines before attempting removal
            if (sparklines.Count > 1)
            {
                try
                {
                    // Delete the second sparkline (index 1) from the group
                    sparklines.RemoveAt(1);
                    Console.WriteLine("Second sparkline removed successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to remove sparkline: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("The sparkline group does not contain a second sparkline to remove.");
            }

            // Save the workbook with the changes
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

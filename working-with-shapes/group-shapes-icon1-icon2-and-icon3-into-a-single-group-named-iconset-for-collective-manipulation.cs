// Title: How to group three named shapes (Icon1, Icon2, Icon3) into a single shape group called IconSet using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that loads a workbook, retrieves shapes named Icon1, Icon2, and Icon3, groups them into a shape named IconSet, and saves the file. | Write a .NET snippet that uses Aspose.Cells to create a shape group from existing worksheet shapes, assign a custom group name, and persist the changes.
// Common Searches: Aspose.Cells C# group multiple shapes into one shape group | How to name a shape group in an Excel file using Aspose.Cells | Retrieve shapes by name and create a group with Aspose.Cells .NET | Save workbook after grouping shapes with Aspose.Cells | C# example for grouping Icon1 Icon2 Icon3 in Excel using Aspose.Cells
// Tags: group shapes Aspose.Cells .NET | shape grouping Excel workbook Aspose.Cells | named shape retrieval Aspose.Cells C# | assign custom name to shape group Aspose.Cells | save workbook after shape manipulation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Loads an existing Excel workbook, fetches shapes named Icon1, Icon2, and Icon3 from the first worksheet, groups them into a new shape called IconSet, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve shapes by their names
                Shape icon1 = worksheet.Shapes["Icon1"];
                Shape icon2 = worksheet.Shapes["Icon2"];
                Shape icon3 = worksheet.Shapes["Icon3"];

                // Ensure all shapes were found before proceeding
                if (icon1 == null || icon2 == null || icon3 == null)
                {
                    Console.WriteLine("Error: One or more specified shapes were not found in the worksheet.");
                    return;
                }

                // Group the shapes
                Shape[] iconsToGroup = new Shape[] { icon1, icon2, icon3 };
                Shape iconSetGroup = worksheet.Shapes.Group(iconsToGroup);
                iconSetGroup.Name = "IconSet";

                // Save the workbook with the new grouped shape
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

// Title: Update a shape's offset to (10,20) in an Excel worksheet using Aspose.Cells for .NET and save the workbook
// AI Prompts: In C#, use Aspose.Cells to add 10 points to the Left property and 20 points to the Top property of the first shape on a worksheet, then save the workbook to a new file. | With Aspose.Cells for .NET, modify a shape's connection point offset to (10,20) and persist the changes by writing the updated Excel file.
// Common Searches: asp.net c# change shape position offset in Excel with Aspose.Cells | increase shape left and top values by 10 and 20 points using Aspose.Cells | save workbook after moving a shape in Excel via Aspose.Cells .NET
// Tags: Aspose.Cells shape offset modification | C# adjust Excel shape coordinates | Aspose.Cells save workbook after shape change | Excel shape left top property update .NET

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an existing Excel file, checks for at least one shape on the first worksheet, adds a 10‑point horizontal and 20‑point vertical offset to the shape's position by updating its Left and Top properties, and then saves the modified workbook to a new file.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one shape
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found in the worksheet.");
                return;
            }

            // Retrieve the target shape (here we use the first shape in the sheet)
            Shape shape = sheet.Shapes[0];

            try
            {
                // Adjust the shape's position by setting its offset.
                // Here we move the shape 10 points to the right and 20 points down.
                shape.Left += 10;
                shape.Top += 20;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error adjusting shape position: {ex.Message}");
            }

            // Persist the changes to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

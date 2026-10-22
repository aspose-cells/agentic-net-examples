// Title: Rename an Excel form control button without breaking its macro using Aspose.Cells for .NET
// Common Searches: C# code to change the name of a form control button in an Excel file with Aspose.Cells | how to keep macro assignment when renaming a button shape in Aspose.Cells workbook | search for a button shape by its name and rename it using Aspose.Cells .NET API | update Excel form control button name programmatically without affecting linked macro | Aspose.Cells example for renaming a button control while preserving macro reference
// Tags: Aspose.Cells rename form control button | preserve macro reference when renaming button | locate Excel shape by name using Aspose.Cells | modify button shape properties .NET | Excel form control manipulation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing Excel workbook, finds a form control button shape by its current name, renames the button while retaining any assigned macro, and saves the updated file.
class RenameFormControlButton
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Define old and new button names
            string oldButtonName = "OldButtonName";
            string newButtonName = "NewButtonName";

            // Get the first worksheet (adjust index as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Find the form control button shape by its current name
            Shape buttonShape = null;
            foreach (Shape shape in sheet.Shapes)
            {
                // Check if the shape has the specified name and is a button control
                if (shape.Name == oldButtonName && shape is Button)
                {
                    buttonShape = shape;
                    break;
                }
            }

            if (buttonShape != null)
            {
                // Rename the button
                buttonShape.Name = newButtonName;

                Console.WriteLine($"Button renamed from '{oldButtonName}' to '{newButtonName}'.");
            }
            else
            {
                Console.WriteLine($"Button with name '{oldButtonName}' not found.");
            }

            // Save the modified workbook
            string outputPath = "output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

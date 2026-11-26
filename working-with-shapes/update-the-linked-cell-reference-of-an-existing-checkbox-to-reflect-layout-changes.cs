// Title: Update the LinkedCell of an existing CheckBox shape in an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells change linked cell of a checkbox after moving it | how to set new LinkedCell address for Excel form control checkbox using Aspose.Cells | update checkbox linked cell programmatically in .NET workbook | Aspose.Cells modify CheckBox LinkedCell property in existing spreadsheet
// Tags: Aspose.Cells CheckBox LinkedCell property | C# update Excel form control linked cell | modify checkbox linked cell address Aspose.Cells | Excel worksheet programmatic checkbox cell reference | Aspose.Cells shape property manipulation C#

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace UpdateCheckBoxLinkedCell
{
    // Loads an existing Excel file, ensures a CheckBox shape is present on the first worksheet, changes its LinkedCell property to a new address (e.g., B2), and saves the workbook with the updated linked cell reference using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output files.
            // Adjust these paths as needed for your environment.
            string inputPath = "input.xlsx";
            string outputPath = "output_updated.xlsx";

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);
                // Assume we are working with the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the collection of check boxes on the worksheet.
                CheckBoxCollection checkBoxes = worksheet.CheckBoxes;

                // If there are no check boxes, create one for demonstration.
                if (checkBoxes.Count == 0)
                {
                    Console.WriteLine("No CheckBox found. Adding a new CheckBox at A1.");
                    int addedIndex = checkBoxes.Add(0, 0, 20, 100); // row, column, height, width
                    CheckBox newBox = checkBoxes[addedIndex];
                    newBox.Text = "Demo CheckBox";
                    // Initially link it to cell A1.
                    newBox.LinkedCell = "A1";
                }

                // For this example, we will update the first CheckBox in the collection.
                CheckBox checkBoxToUpdate = checkBoxes[0];

                // Display current linked cell.
                Console.WriteLine($"Current LinkedCell of CheckBox[0]: {checkBoxToUpdate.LinkedCell}");

                // Define the new linked cell address.
                // This could be calculated based on layout changes; here we use a static example.
                string newLinkedCell = "B2";

                // Update the LinkedCell property.
                checkBoxToUpdate.LinkedCell = newLinkedCell;

                // Verify the update.
                Console.WriteLine($"LinkedCell updated to: {checkBoxToUpdate.LinkedCell}");

                // Save the workbook with the updated linked cell reference.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Output any errors that occur during processing.
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}

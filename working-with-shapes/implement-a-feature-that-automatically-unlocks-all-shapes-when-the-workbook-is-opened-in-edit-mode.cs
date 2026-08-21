// Title: C# – Unlock All Shapes on Workbook Open with Aspose.Cells
// Description: Loads a workbook, enables editing of objects on protected sheets (AllowEditingObject = true), sets each shape’s IsLocked property to false, and saves the file so all shapes are editable when the workbook is opened.
// Keywords: Aspose.Cells C# unlock shapes | worksheet protection allow editing objects | shape.IsLocked false | unlock all shapes Excel | auto unlock shapes on open | Aspose.Cells shape unlocking | C# Excel shape editing | protect sheet keep shapes unlocked
// Common Searches: unlock all shapes in protected Excel sheet using Aspose.Cells | Aspose.Cells set AllowEditingObject true | C# code to make shapes editable after opening workbook | how to programmatically unlock shapes in Excel with Aspose | auto unlock shapes on workbook open .NET
// Developer Intent: Programmatically ensure every shape in an Excel workbook is unlocked and editable even when the worksheet remains protected.
// Use Cases: Distribute a template where users can modify charts, images, or buttons immediately after opening the file. | Process incoming workbooks from partners, remove shape locks, and re‑save them for downstream editing pipelines. | Integrate into an automated report generator that must keep all inserted shapes editable without disabling sheet protection.
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, sets worksheet.Protection.AllowEditingObject = true, iterates through worksheet.Shapes, sets shape.IsLocked = false, and saves the workbook. | Explain best practices for handling missing input files and exceptions while unlocking all shapes in a protected workbook using Aspose.Cells. | Show how to unlock shapes in a workbook loaded from a MemoryStream instead of a file path. | Provide a step‑by‑step guide to keep sheet protection enabled while making all shapes editable in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Loads a workbook, enables editing of objects on protected sheets (AllowEditingObject = true), sets each shape’s IsLocked property to false, and saves the file so all shapes are editable when the workbook is opened.
    public class UnlockAllShapesOnOpen
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet in the workbook
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    // Allow editing of objects on a protected sheet (does not remove protection)
                    worksheet.Protection.AllowEditingObject = true;

                    // Unlock each shape so it can be modified even when the sheet is protected
                    foreach (Shape shape in worksheet.Shapes)
                    {
                        shape.IsLocked = false;
                    }
                }

                // Save the workbook after unlocking all shapes
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}

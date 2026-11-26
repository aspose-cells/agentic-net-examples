// Title: Unlock a locked textbox shape and modify its position and size using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# unlock textbox shape on protected worksheet | change textbox dimensions after unlocking shape in Aspose.Cells | C# code to modify position of a locked textbox in Excel using Aspose.Cells | unprotect worksheet and edit textbox size with Aspose.Cells for .NET | set IsLocked false for textbox shape in Aspose.Cells example
// Tags: textbox shape IsLocked flag Aspose.Cells | shape Left Top Width Height adjustment Aspose.Cells | worksheet protection handling Aspose.Cells C# | shape dimension editing Aspose.Cells .NET | unlock and move textbox Aspose.Cells sample

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace UnlockTextboxDemo
{
    // The example loads an Excel workbook, optionally unprotects the first worksheet, iterates through its shapes to locate textbox AutoShape objects, clears the IsLocked flag, and then sets new Left, Top, Width, and Height values before saving the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the input and output Excel files.
            string inputFile = "LockedTextbox.xlsx";
            string outputFile = "UnlockedTextbox.xlsx";

            try
            {
                // Verify that the input file exists.
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file \"{inputFile}\" not found.");
                    return;
                }

                // Load the workbook that contains the locked textbox.
                Workbook workbook = new Workbook(inputFile);

                // Work with the first worksheet (adjust index if needed).
                Worksheet worksheet = workbook.Worksheets[0];

                // If the worksheet is protected, unprotect it first.
                // This is required because a protected sheet can prevent shape modifications.
                if (worksheet.IsProtected)
                {
                    // No password assumed; pass null or empty string if a password was set.
                    worksheet.Unprotect();
                }

                // Iterate through all shapes on the worksheet.
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Identify textbox shapes.
                    if (shape.AutoShapeType == AutoShapeType.TextBox)
                    {
                        Console.WriteLine($"Found TextBox shape: Name=\"{shape.Name}\"");

                        // Unlock the shape so its position and size can be edited.
                        shape.IsLocked = false;

                        // Example: change its position and size to demonstrate that it is now editable.
                        shape.Left = 100;   // distance from the left edge of the worksheet
                        shape.Top = 50;     // distance from the top edge of the worksheet
                        shape.Width = 300;  // width of the textbox
                        shape.Height = 100; // height of the textbox

                        Console.WriteLine("Textbox unlocked and repositioned.");
                    }
                }

                // Save the modified workbook.
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved to \"{outputFile}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}

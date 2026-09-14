// Title: Unlock the 'Signature' shape in an Excel workbook and resize it with Aspose.Cells for .NET
// AI Prompts: Write C# code that searches a worksheet for a shape named "Signature", asks the user to confirm, sets IsLocked = false, and updates its Width, Height, Left, and Top properties using Aspose.Cells. | Generate a console application that loads an Excel file, locates a locked shape, unlocks it, resizes it to given dimensions, and saves the modified workbook with Aspose.Cells for .NET. | Provide a step‑by‑step script that prompts the user before unlocking a shape and then programmatically changes the shape's size and position in a workbook using Aspose.Cells.
// Common Searches: Aspose.Cells how to unlock a specific shape in C# | C# resize Excel shape after unlocking with Aspose.Cells | programmatically change position of a named shape in an Excel workbook using Aspose.Cells | prompt user before modifying locked shape Aspose.Cells .NET
// Tags: unlock Aspose.Cells shape programmatically | resize shape dimensions Aspose.Cells C# | set shape position Aspose.Cells workbook | prompt user before shape modification .NET | find shape by name Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads input.xlsx, finds the shape named "Signature" on the first worksheet, prompts the user for confirmation, unlocks the shape by setting IsLocked to false, changes its Width, Height, Left, and Top values, and saves the updated workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Locate the shape named "Signature"
            Shape signatureShape = null;
            foreach (Shape shape in sheet.Shapes)
            {
                if (shape.Name == "Signature")
                {
                    signatureShape = shape;
                    break;
                }
            }

            if (signatureShape == null)
            {
                Console.WriteLine("Signature shape not found.");
                return;
            }

            // Ask the user for confirmation before unlocking
            Console.Write("Unlock the signature shape to allow size/position changes? (y/n): ");
            string answer = Console.ReadLine();
            if (!string.Equals(answer, "y", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("Operation cancelled by user.");
                return;
            }

            // Unlock the shape (Aspose.Cells does not expose a LockAspectRatio property)
            signatureShape.IsLocked = false;

            // Example modifications – adjust as needed (values are in points)
            signatureShape.Width = 200;
            signatureShape.Height = 100;
            signatureShape.Left = 50;
            signatureShape.Top = 30;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

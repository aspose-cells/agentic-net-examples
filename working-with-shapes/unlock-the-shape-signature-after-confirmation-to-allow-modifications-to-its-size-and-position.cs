// Title: Unlock a locked signature shape and modify its size and position using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# unlock shape IsLocked property before resizing | Change position of a protected shape in Excel with Aspose.Cells | How to unlock and move a signature rectangle shape in a worksheet using Aspose.Cells | C# example to modify locked shape dimensions after removing worksheet protection | Prompt user to confirm unlocking of a shape in Aspose.Cells code
// Tags: Aspose.Cells unlock shape IsLocked | Aspose.Cells modify shape width height | Aspose.Cells adjust shape left top | Aspose.Cells worksheet objects protection | Aspose.Cells signature placeholder shape

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace UnlockSignatureShapeDemo
{
    // The sample loads or creates an Excel workbook, adds a rectangle named "Signature", protects the worksheet for objects, then unprotects the sheet, prompts the user for confirmation, sets the shape's IsLocked flag to false, updates its Width, Height, Left, and Top properties, and saves the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output workbooks.
            string inputPath = "SignatureDemo.xlsx";
            string outputPath = "SignatureUnlocked.xlsx";

            Workbook workbook = null;

            try
            {
                // Load existing workbook or create a sample one if it does not exist.
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook \"{inputPath}\".");
                }
                else
                {
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];
                    ws.Name = "Sheet1";

                    // Add a rectangle shape to act as a signature placeholder.
                    Shape signatureShape = ws.Shapes.AddAutoShape(
                        AutoShapeType.Rectangle,
                        2,   // upper left row
                        1,   // upper left column
                        0,   // row offset (pixels)
                        0,   // column offset (pixels)
                        150, // width (points)
                        50   // height (points)
                    );
                    signatureShape.Name = "Signature";
                    signatureShape.Text = "Signed by: ____________________";

                    // Set fill color using the correct API.
                    signatureShape.Fill.SolidFill.Color = System.Drawing.Color.LightGray;

                    signatureShape.Line.Weight = 1;
                    signatureShape.Line.DashStyle = MsoLineDashStyle.Solid;

                    // Protect the worksheet for objects so the shape becomes locked.
                    ws.Protect(ProtectionType.Objects);
                    Console.WriteLine("Created sample workbook with a locked signature shape.");

                    workbook.Save(inputPath);
                    Console.WriteLine($"Sample workbook saved as \"{inputPath}\".");
                }

                // Access the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];

                // Unprotect the worksheet if it is protected.
                try
                {
                    worksheet.Unprotect();
                    Console.WriteLine("Worksheet unprotected to allow shape editing.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to unprotect worksheet: {ex.Message}");
                }

                // Locate the shape named "Signature".
                Shape signature = null;
                foreach (Shape shape in worksheet.Shapes)
                {
                    if (shape.Name.Equals("Signature", StringComparison.OrdinalIgnoreCase))
                    {
                        signature = shape;
                        break;
                    }
                }

                if (signature == null)
                {
                    Console.WriteLine("No shape named \"Signature\" was found in the worksheet.");
                    return;
                }

                // Ask user for confirmation before unlocking.
                Console.Write("Do you want to unlock the signature shape for editing? (y/n): ");
                string answer = Console.ReadLine()?.Trim().ToLowerInvariant();

                if (answer != "y" && answer != "yes")
                {
                    Console.WriteLine("Operation cancelled by the user.");
                    return;
                }

                // Unlock the shape.
                signature.IsLocked = false;
                Console.WriteLine("Signature shape unlocked.");

                // Modify size and position (units are points).
                signature.Width = 300;
                signature.Height = 80;
                signature.Left = 2000;
                signature.Top = 1500;

                Console.WriteLine("Signature shape size and position have been updated.");
                Console.WriteLine($"New Width: {signature.Width}, Height: {signature.Height}");
                Console.WriteLine($"New Left: {signature.Left}, Top: {signature.Top}");

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Modified workbook saved as \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

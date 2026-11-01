// Title: Remove a specific ActiveX control from an Excel worksheet and confirm its removal using Aspose.Cells for .NET
// Common Searches: how to delete an ActiveX control from an Excel worksheet using Aspose.Cells C# | Aspose.Cells remove shape by name and verify it no longer exists | C# code to remove a checkbox ActiveX control from a workbook and check Shapes collection | remove specific ActiveX control from worksheet and save workbook with Aspose.Cells | verify that an ActiveX control has been removed from Excel using Aspose.Cells API
// Tags: delete ActiveX shape by name Aspose.Cells | verify shape absence in worksheet Aspose.Cells | remove checkbox control from Excel using Aspose.Cells | save workbook after shape removal Aspose.Cells | iterate Shapes collection to find control Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel file, searches the first worksheet's Shapes collection for a shape named "CheckBox1", removes that ActiveX control with Shapes.RemoveAt, then iterates the collection again to confirm the control is no longer present before saving the modified workbook to a new file.
class RemoveActiveXControl
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "InputWorkbook.xlsx";
            string outputPath = "OutputWorkbook.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or specify by name/index)
            Worksheet worksheet = workbook.Worksheets[0];

            // Name of the ActiveX control to remove (adjust as needed)
            string targetControlName = "CheckBox1";

            // Find the index of the shape with the specified name
            int shapeIndexToRemove = -1;
            for (int i = 0; i < worksheet.Shapes.Count; i++)
            {
                Shape shape = worksheet.Shapes[i];
                if (shape.Name == targetControlName)
                {
                    shapeIndexToRemove = i;
                    break;
                }
            }

            // Remove the shape if it was found
            if (shapeIndexToRemove != -1)
            {
                worksheet.Shapes.RemoveAt(shapeIndexToRemove);
                Console.WriteLine($"ActiveX control \"{targetControlName}\" removed.");
            }
            else
            {
                Console.WriteLine($"ActiveX control \"{targetControlName}\" not found.");
            }

            // Verify that the control is no longer present
            bool stillExists = false;
            foreach (Shape shape in worksheet.Shapes)
            {
                if (shape.Name == targetControlName)
                {
                    stillExists = true;
                    break;
                }
            }

            Console.WriteLine(stillExists
                ? $"Verification failed: \"{targetControlName}\" still exists."
                : $"Verification succeeded: \"{targetControlName}\" is absent.");

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

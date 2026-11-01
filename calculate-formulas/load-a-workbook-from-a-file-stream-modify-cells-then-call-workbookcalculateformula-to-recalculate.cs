// Title: Load an Excel workbook from a FileStream, edit cells, set a SUM formula, and recalculate all formulas with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load workbook from FileStream and recalculate formulas after editing cells | How to modify cell values and set a SUM formula using Aspose.Cells before saving | Recalculate all formulas in an Excel file after updating data with Aspose.Cells .NET | Save an Aspose.Cells workbook to a new file stream after changing cells | Create a blank workbook when the source file is missing and process it with Aspose.Cells
// Tags: load workbook via filestream Aspose.Cells | edit cells and assign formula Aspose.Cells C# | Workbook.CalculateFormula method usage | save workbook to output stream Aspose.Cells | initialize workbook if source file absent Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for an input.xlsx file, creates a blank workbook if it doesn't exist, loads the workbook from a FileStream, updates cells A1 and B2, assigns a SUM formula to C3, recalculates all formulas with Workbook.CalculateFormula, and saves the result to output.xlsx using another FileStream.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; create a blank workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                tempWb.Save(inputPath, SaveFormat.Xlsx);
            }

            // Load workbook from a file stream
            using (FileStream inputStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            {
                // Create Workbook instance from the stream
                Workbook workbook = new Workbook(inputStream);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Modify some cells
                sheet.Cells["A1"].PutValue("Hello World");   // Text value
                sheet.Cells["B2"].PutValue(12345);           // Numeric value

                // Set a formula (use the Formula property for compatibility)
                sheet.Cells["C3"].Formula = "=SUM(B2:B10)";

                // Recalculate all formulas in the workbook
                workbook.CalculateFormula();

                // Save the modified workbook to a new file stream
                using (FileStream outputStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Save(outputStream, SaveFormat.Xlsx);
                }
            }

            Console.WriteLine("Workbook processed and saved to '{0}'.", outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}

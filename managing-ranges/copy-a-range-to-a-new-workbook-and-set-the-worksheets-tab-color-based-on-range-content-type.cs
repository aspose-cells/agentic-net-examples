// Title: Copy a cell range to a new workbook and set the worksheet tab color according to the range’s data type using Aspose.Cells for .NET
// AI Prompts: Copy the A1:C10 range from an existing workbook into a fresh workbook while preserving values with Aspose.Cells in C#. | Iterate through the copied range, identify whether each cell contains a numeric, string, or DateTime value, and assign the destination worksheet’s TabColor to LightBlue, LightGreen, LightCoral, or LightGray based on the detected content type. | Save the resulting workbook as output.xlsx after the range copy and conditional tab‑coloring steps.
// Common Searches: Aspose.Cells copy range to another workbook and change worksheet tab color based on cell type | C# detect numeric, text, and date cells in a copied range using Aspose.Cells | set worksheet TabColor conditionally after copying a range in Aspose.Cells .NET | copy range A1:C10 to new workbook and apply tab color coding with Aspose.Cells
// Tags: copy range to new workbook Aspose.Cells | conditional tab color based on cell type Aspose.Cells | detect numeric string date cells C# Aspose.Cells | set worksheet TabColor programmatically .NET | range content analysis Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads source.xlsx, copies the A1:C10 range to a new workbook, examines each cell to determine if the data is numeric, text, or date, sets the destination worksheet's TabColor (LightBlue, LightGreen, LightCoral, or LightGray) according to the detected content type, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string outputPath = "output.xlsx";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file \"{sourcePath}\" not found.");
                return;
            }

            // Load the source workbook
            Workbook srcWb = new Workbook(sourcePath);

            // Create a new (empty) workbook
            Workbook destWb = new Workbook();

            // Get the first worksheets from both workbooks
            Worksheet srcSheet = srcWb.Worksheets[0];
            Worksheet destSheet = destWb.Worksheets[0];

            // Define the range to copy (adjust as needed)
            Aspose.Cells.Range srcRange = srcSheet.Cells.CreateRange("A1:C10");

            // Copy the range to the destination worksheet starting at A1
            destSheet.Cells.CreateRange("A1").Copy(srcRange);

            // Analyze the content type of the copied range
            bool hasNumeric = false;
            bool hasString = false;
            bool hasDate = false;

            foreach (Cell cell in srcRange)
            {
                if (cell.Type == CellValueType.IsNumeric)
                    hasNumeric = true;
                else if (cell.Type == CellValueType.IsString)
                    hasString = true;
                else if (cell.Type == CellValueType.IsDateTime)
                    hasDate = true;
            }

            // Set the worksheet tab color based on the detected content type
            if (hasNumeric && !hasString && !hasDate)
                destSheet.TabColor = Color.LightBlue;      // Only numeric values
            else if (hasString && !hasNumeric && !hasDate)
                destSheet.TabColor = Color.LightGreen;     // Only text values
            else if (hasDate && !hasNumeric && !hasString)
                destSheet.TabColor = Color.LightCoral;     // Only date values
            else
                destSheet.TabColor = Color.LightGray;      // Mixed content

            // Save the new workbook
            destWb.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

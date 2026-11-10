// Title: Create a non‑contiguous UnionRange (M1:M5, O1:O5) and apply a uniform two‑decimal number format using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# create union range with multiple areas M1:M5 and O1:O5 | Apply same number format to non‑adjacent cells in Aspose.Cells .NET | Set custom numeric style 0.00 for a union range in a workbook using Aspose.Cells | Use StyleFlag.All to format all cells in a union range with Aspose.Cells | Save workbook after formatting union range in Aspose.Cells C#
// Tags: union range creation Aspose.Cells .NET | uniform numeric format Aspose.Cells | styleflag all Aspose.Cells | non‑adjacent cell formatting Aspose.Cells | save formatted workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, defines a UnionRange covering cells M1:M5 and O1:O5, builds a style with the custom number format "0.00", applies this style to the entire union range using a StyleFlag with All set to true, and saves the workbook as UnionRangeFormatted.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a UnionRange that includes cells M1:M5 and O1:O5
            Aspose.Cells.Range unionRange = worksheet.Cells.CreateRange("M1:M5,O1:O5");

            // Define a uniform number format (two decimal places)
            Style numberStyle = workbook.CreateStyle();
            numberStyle.Custom = "0.00";

            // Apply the number format to all cells in the union range
            // Use a StyleFlag that applies all style attributes
            StyleFlag flag = new StyleFlag { All = true };
            unionRange.ApplyStyle(numberStyle, flag);

            // Determine output path and ensure directory exists
            string outputFile = "UnionRangeFormatted.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));

            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

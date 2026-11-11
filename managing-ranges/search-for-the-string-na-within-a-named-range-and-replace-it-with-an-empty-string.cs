// Title: How to replace "#N/A" error text with an empty string in a named range using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# replace #N/A error value in a named range | remove #N/A strings from specific Excel named range with Aspose.Cells | iterate over named range cells and clear #N/A using Aspose.Cells .NET | C# code to find and delete #N/A entries in a workbook named range | Aspose.Cells find exact string "#N/A" in range and set blank
// Tags: replace #N/A in named range Aspose.Cells | clear error string in Excel range C# | Aspose.Cells iterate named range cells | set empty value for specific cell content Aspose.Cells | Excel workbook named range manipulation Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook (or creates a new one), obtains the named range "MyRange", walks through each cell in that range, and replaces any cell whose string value is exactly "#N/A" with an empty string, then saves the result to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load workbook if the input file exists; otherwise create a new workbook
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Retrieve the range associated with the named range "MyRange"
            Aspose.Cells.Range namedRange = workbook.Worksheets.GetRangeByName("MyRange");

            if (namedRange != null)
            {
                // Iterate through each cell in the range
                foreach (Cell cell in namedRange)
                {
                    // Replace exact "#N/A" string with an empty string
                    if (cell.Type == CellValueType.IsString && cell.StringValue == "#N/A")
                    {
                        cell.PutValue(string.Empty);
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

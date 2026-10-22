// Title: Copy a range from one Excel workbook to a new workbook and color the worksheet tab according to the dominant cell data type using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy specific range to a new workbook and set worksheet tab color | How to change Excel worksheet tab color based on the most common cell type with Aspose.Cells | Determine predominant data type in a range and apply tab coloring using Aspose.Cells .NET | Copy A1:C10 to another workbook and color the tab by numeric, text, date, or boolean values Aspose | C# example for range copy and conditional worksheet tab coloring with Aspose.Cells
// Tags: duplicate range into new workbook Aspose.Cells | worksheet tab color based on dominant cell type | determine most frequent cell value type in Excel range | conditional tab styling using Aspose.Cells C# | C# analysis of cell data types for tab coloring

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The sample loads a source Excel file, copies a defined range (e.g., A1:C10) into a freshly created workbook, scans the source range to count numeric, text, date, boolean and other cell types, selects the most frequent type, sets the destination worksheet's tab color accordingly (numeric → LightBlue, text → LightGreen, date → LightCoral, boolean → LightYellow, mixed/other → LightGray), and saves the new workbook.
class Program
{
    static void Main()
    {
        const string sourcePath = "source.xlsx";
        const string destPath = "output.xlsx";

        // Verify source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook sourceWb = new Workbook(sourcePath);

            // Define the range to copy
            string sourceRangeAddress = "A1:C10";
            Worksheet sourceWs = sourceWb.Worksheets[0];
            AsposeRange sourceRange = sourceWs.Cells.CreateRange(sourceRangeAddress);

            // Create a new workbook and add a fresh worksheet
            Workbook destWb = new Workbook();
            destWb.Worksheets.Clear();
            Worksheet destWs = destWb.Worksheets.Add("CopiedRange");

            // Copy the range to the destination worksheet starting at A1 (row 0, column 0)
            AsposeRange destRange = destWs.Cells.CreateRange(
                0,                     // first row (A)
                0,                     // first column (1)
                sourceRange.RowCount,
                sourceRange.ColumnCount);
            destRange.Copy(sourceRange);

            // Determine the predominant content type in the source range
            int numericCount = 0, textCount = 0, dateCount = 0, boolCount = 0, otherCount = 0;

            foreach (Cell cell in sourceRange)
            {
                // Skip empty cells
                if (cell.Type == CellValueType.IsNull)
                    continue;

                switch (cell.Type)
                {
                    case CellValueType.IsNumeric:
                        numericCount++;
                        break;
                    case CellValueType.IsString:
                        textCount++;
                        break;
                    case CellValueType.IsDateTime:
                        dateCount++;
                        break;
                    case CellValueType.IsBool:
                        boolCount++;
                        break;
                    default:
                        otherCount++;
                        break;
                }
            }

            // Choose the tab color based on the most frequent content type
            if (numericCount >= textCount && numericCount >= dateCount && numericCount >= boolCount && numericCount >= otherCount)
                destWs.TabColor = Color.LightBlue;      // Numeric data
            else if (textCount >= numericCount && textCount >= dateCount && textCount >= boolCount && textCount >= otherCount)
                destWs.TabColor = Color.LightGreen;     // Text data
            else if (dateCount >= numericCount && dateCount >= textCount && dateCount >= boolCount && dateCount >= otherCount)
                destWs.TabColor = Color.LightCoral;     // Date data
            else if (boolCount >= numericCount && boolCount >= textCount && boolCount >= dateCount && boolCount >= otherCount)
                destWs.TabColor = Color.LightYellow;    // Boolean data
            else
                destWs.TabColor = Color.LightGray;      // Mixed/Other

            // Save the destination workbook
            destWb.Save(destPath);
            Console.WriteLine($"Workbook saved successfully to {destPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

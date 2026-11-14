// Title: Convert comma‑separated strings in column B to separate columns using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# split comma‑separated values in a column into multiple cells | How to parse CSV text in Excel column B and write to adjacent columns with Aspose.Cells | C# code to separate comma‑delimited strings in an Excel worksheet using Aspose.Cells | Iterate rows and split cell values by comma delimiter in Aspose.Cells .NET | Save Excel after expanding CSV values from one column to several columns with Aspose.Cells
// Tags: csv column splitting Aspose.Cells | comma‑delimited string parsing Excel .NET | populate adjacent cells from split values Aspose.Cells | row iteration and cell splitting Aspose.Cells | excel workbook save after column expansion .NET

using Aspose.Cells;
using System;

// The example loads an Excel workbook, walks through each used row in column B, splits any non‑empty comma‑separated string, writes each trimmed element into consecutive cells starting at column C, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Determine the last row that contains data
        int lastRow = cells.MaxDataRow;

        // Iterate through each row in column B (index 1)
        for (int row = 0; row <= lastRow; row++)
        {
            Cell sourceCell = cells[row, 1]; // Column B

            // Skip empty cells
            if (sourceCell.Type == CellValueType.IsNull)
                continue;

            string csv = sourceCell.StringValue;
            if (string.IsNullOrWhiteSpace(csv))
                continue;

            // Split the CSV string by comma
            string[] parts = csv.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            // Write each part into subsequent columns starting from column C (index 2)
            for (int i = 0; i < parts.Length; i++)
            {
                string value = parts[i].Trim(); // Remove surrounding spaces
                cells[row, 2 + i].PutValue(value);
            }
        }

        // Save the modified workbook (replace with your desired output path)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}

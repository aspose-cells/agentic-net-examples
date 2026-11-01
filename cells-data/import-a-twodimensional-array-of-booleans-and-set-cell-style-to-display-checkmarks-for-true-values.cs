// Title: Write a 2D bool[,] array to an Excel worksheet and display checkmarks for true values using Aspose.Cells for .NET
// AI Prompts: Generate C# code that iterates over a bool[,] matrix, writes a checkmark character to cells with true values, and leaves false cells blank using Aspose.Cells. | Create a Wingdings 2 style in Aspose.Cells and apply it only to cells that represent true entries from a boolean array. | Save the populated workbook as an .xlsx file after inserting checkmark symbols for true values.
// Common Searches: Aspose.Cells C# write boolean matrix to Excel with checkmark symbols | How to display a checkmark for true values in an Excel file using Aspose.Cells | Create custom cell style with Wingdings 2 font in Aspose.Cells .NET | Import two dimensional bool array into worksheet using Aspose.Cells | Set cell value to 'P' for checkmark in Aspose.Cells example
// Tags: Aspose.Cells write boolean array to worksheet | Wingdings 2 checkmark style Aspose.Cells | C# bool[,] to Excel with custom font | apply custom style to true cells Aspose.Cells | export boolean matrix as checkmarks .xlsx

using Aspose.Cells;
using System.Drawing;

// Sample two‑dimensional boolean array
bool[,] boolArray = new bool[,]
{
    { true, false, true },
    { false, true, false },
    { true, true, false }
};

// Create a new workbook (Aspose.Cells)
Workbook workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];

// Create a style that uses a font capable of displaying a checkmark (Wingdings 2)
Style checkStyle = workbook.CreateStyle();
checkStyle.Font.Name = "Wingdings 2";
checkStyle.Font.Size = 12;
checkStyle.Font.Color = Color.Black;
checkStyle.HorizontalAlignment = TextAlignmentType.Center;
checkStyle.VerticalAlignment = TextAlignmentType.Center;

// Iterate over the boolean array and write values to cells
for (int row = 0; row < boolArray.GetLength(0); row++)
{
    for (int col = 0; col < boolArray.GetLength(1); col++)
    {
        Cell cell = sheet.Cells[row, col];

        if (boolArray[row, col])
        {
            // In Wingdings 2 the character 'P' renders as a checkmark
            cell.PutValue("P");
            cell.SetStyle(checkStyle);
        }
        else
        {
            cell.PutValue(string.Empty);
        }
    }
}

// Save the workbook to a file
workbook.Save("BooleanCheckmarks.xlsx");

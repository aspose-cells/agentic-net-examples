// Title: Import a 2‑D boolean array into an Excel sheet and show true values as checkmarks with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# import bool[,] into worksheet and format as checkmark | display checkmark symbol for true booleans using Aspose.Cells | apply custom number format to a boolean range in an Aspose.Cells workbook | convert two‑dimensional bool array to object[,] for ImportTwoDimensionArray | use StyleFlag to apply only number format in Aspose.Cells C#
// Tags: ImportTwoDimensionArray boolean array Aspose.Cells | custom number format checkmark Aspose.Cells | StyleFlag number format only Aspose.Cells | bool[,] to object[,] conversion Aspose.Cells | range.ApplyStyle checkmark true values

using System;
using Aspose.Cells;

// The example creates a 2‑D bool array, converts it to an object[,] required by ImportTwoDimensionArray, imports the data into the first worksheet, defines a style with a custom number format "✔;" that renders true values as checkmarks, applies this style to the imported range using a StyleFlag that limits the change to the number format, and saves the workbook as BooleanCheckmarks.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Sample two‑dimensional boolean array
            bool[,] boolData = new bool[,]
            {
                { true, false, true },
                { false, true, false },
                { true, true, false }
            };

            // Convert bool[,] to object[,] because ImportTwoDimensionArray expects object[,]
            int rows = boolData.GetLength(0);
            int cols = boolData.GetLength(1);
            object[,] objData = new object[rows, cols];
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                    objData[i, j] = boolData[i, j];

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();                     // create rule
            Worksheet sheet = workbook.Worksheets[0];

            // Import the boolean array starting at cell A1 (row 0, column 0)
            sheet.Cells.ImportTwoDimensionArray(objData, 0, 0);    // import rule

            // Create a style that shows a checkmark (✔) for true values
            Style checkStyle = workbook.CreateStyle();             // create rule
            // Custom number format: display a checkmark for positive numbers (true = 1), blank otherwise
            checkStyle.Custom = "✔;";                               // number format rule

            // Apply the style to the entire imported range
            Aspose.Cells.Range range = sheet.Cells.CreateRange(0, 0, rows, cols);
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;      // apply only the number format part of the style
            range.ApplyStyle(checkStyle, flag);

            // Save the workbook to a file
            workbook.Save("BooleanCheckmarks.xlsx");               // save rule
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

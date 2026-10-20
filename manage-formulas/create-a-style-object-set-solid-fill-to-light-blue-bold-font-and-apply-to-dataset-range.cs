// Title: Apply a light‑blue solid fill and bold font style to a DataSet range using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# apply solid fill color to a range of cells | How to set bold font and background color for an imported DataSet in Aspose.Cells | Create and apply custom style to a worksheet range with Aspose.Cells .NET | Use StyleFlag to style a DataSet range in Aspose.Cells C# | Set light blue background for cells using Aspose.Cells style object
// Tags: Aspose.Cells create style with solid fill | apply custom style to worksheet range | set bold font and background color Aspose.Cells | StyleFlag All attribute for range styling | style DataSet range in Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;

// Demonstrates how to create a Style with a light‑blue solid fill and bold font, configure a StyleFlag to include all attributes, and apply the style to a specified cell range representing a DataSet in an Aspose.Cells workbook before saving the file as StyledData.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // OPTIONAL: Import a DataSet into the worksheet.
            // Replace this placeholder with actual DataSet code.
            // DataSet ds = GetYourDataSet();
            // sheet.Cells.ImportDataTable(ds.Tables[0], true, "A1");
            // -------------------------------------------------

            // Define the range that contains the data set.
            // Adjust these values to match the actual size of your data.
            int firstRow = 0;          // Row index for the first row (0 = A1)
            int firstColumn = 0;       // Column index for the first column (0 = A)
            int totalRows = 10;        // Number of rows in the data set
            int totalColumns = 4;      // Number of columns in the data set

            // Create a new style object
            Style style = workbook.CreateStyle();

            // Set solid fill to light blue
            style.ForegroundColor = Color.LightBlue;
            style.Pattern = BackgroundType.Solid;

            // Set bold font
            style.Font.IsBold = true;

            // Prepare a StyleFlag to apply all style attributes
            StyleFlag flag = new StyleFlag { All = true };

            // Apply the style to the defined range
            Aspose.Cells.Range range = sheet.Cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns);
            range.ApplyStyle(style, flag);

            // Save the workbook to a file
            workbook.Save("StyledData.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

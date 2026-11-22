// Title: Apply a mixed‑fraction custom number format (# ?/?) to Excel cells with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set custom number format to display mixed fractions | How to use "# ?/?" format string in Aspose.Cells for .NET | Formatting decimal values as mixed numbers in an Excel workbook using Aspose.Cells | C# code example for applying mixed‑fraction number format to cells with Aspose.Cells | Save Excel file with fractions displayed as mixed numbers using Aspose.Cells
// Tags: mixed fraction custom format Aspose.Cells C# | set cell number format Aspose.Cells .NET | apply custom numeric style to Excel worksheet | display fractions as mixed numbers using Aspose.Cells | custom numeric format for fractions in Aspose.Cells

using Aspose.Cells;

// The example creates a new workbook, writes several decimal values to the first column, applies the custom mixed‑fraction format "# ?/?" to each cell so the numbers appear as mixed fractions, and saves the file as MixedFractions.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Access the first worksheet
        var sheet = workbook.Worksheets[0];

        // Sample numeric values to display as mixed fractions
        double[] values = { 1.25, 2.5, 3.75, 4.125, 5.6667 };

        // Write values to cells and apply the mixed‑fraction custom format
        for (int i = 0; i < values.Length; i++)
        {
            var cell = sheet.Cells[i, 0];
            cell.PutValue(values[i]);

            // Create a style with the custom mixed‑fraction format "# ?/?"
            var style = cell.GetStyle();
            style.Custom = "# ?/?";
            cell.SetStyle(style);
        }

        // Save the workbook
        workbook.Save("MixedFractions.xlsx");
    }
}

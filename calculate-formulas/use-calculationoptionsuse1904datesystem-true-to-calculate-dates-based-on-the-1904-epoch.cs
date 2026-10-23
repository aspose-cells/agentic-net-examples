// Title: Enable the 1904 date system in Aspose.Cells for .NET to calculate formulas using the 1904 epoch (C#)
// Common Searches: Aspose.Cells C# set workbook to use 1904 date system for formula evaluation | How to calculate Excel dates with the 1904 epoch using Aspose.Cells .NET | Enable Use1904DateSystem in Aspose.Cells workbook settings via reflection | Recalculate formulas after switching to the 1904 date system in Aspose.Cells C# | Saving a workbook with the 1904 date system using Aspose.Cells for .NET
// Tags: Aspose.Cells workbook Use1904DateSystem configuration | C# Excel 1904 epoch formula calculation | Aspose.Cells date system compatibility .NET | Reflection setting Use1904DateSystem Aspose.Cells | Excel 1904 date system support in Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new Workbook, uses reflection to set Settings.Use1904DateSystem to true, inserts a DateTime value, recalculates all formulas, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable the 1904 date system if the property exists (compatible with various Aspose.Cells versions)
            var use1904Prop = workbook.Settings?.GetType().GetProperty("Use1904DateSystem");
            if (use1904Prop != null && use1904Prop.CanWrite)
            {
                use1904Prop.SetValue(workbook.Settings, true);
            }

            // Put a date value into a cell
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(new DateTime(2020, 1, 1));

            // Recalculate all formulas (default calculation options are sufficient)
            workbook.CalculateFormula();

            // Save the workbook to a file
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

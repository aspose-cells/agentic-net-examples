// Title: Aspose.Cells .NET – Add a Line Sparkline (Excel 2010) and Note Soft‑Lighting Limitation
// Description: C# example that creates a workbook, inserts a line‑type sparkline for range A1:D1, places it in cell E1, and saves as an Excel 2010 .xlsx file. The API does not expose a property for soft lighting or 3‑D effects, so visual tweaks must rely on available sparkline styling options.
// Keywords: Aspose.Cells | C# sparkline example | Excel 2010 sparkline | soft lighting not supported | sparkline styling | SparklineGroup | line sparkline | visual appearance | Aspose.Cells API limitation
// Common Searches: Aspose.Cells sparkline soft lighting | how to change sparkline lighting in .NET | sparkline visual style Aspose.Cells | apply 3D effects to sparklines with Aspose | sparkline appearance options Excel 2010
// Developer Intent: Add a line sparkline to an Excel 2010 file and understand that soft lighting cannot be set via Aspose.Cells.
// Use Cases: Generate a workbook and insert a line sparkline for a data range. | Attempt to modify sparkline lighting (discover API limitation). | Apply alternative styling such as color, weight, or markers to achieve a softer look. | Save the result as an .xlsx file compatible with Excel 2010.
// AI Prompts: Provide C# code using Aspose.Cells to create a line sparkline and describe how to mimic soft lighting with existing properties. | Suggest a workaround for achieving a softer visual effect on sparklines when the API lacks a lighting setting. | Explain why Aspose.Cells does not support sparkline lighting adjustments and list alternative customization options.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineLightingDemo
{
    // C# example that creates a workbook, inserts a line‑type sparkline for range A1:D1, places it in cell E1, and saves as an Excel 2010 .xlsx file. The API does not expose a property for soft lighting or 3‑D effects, so visual tweaks must rely on available sparkline styling options.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the sparkline
                sheet.Cells["A1"].PutValue(5);
                sheet.Cells["B1"].PutValue(2);
                sheet.Cells["C1"].PutValue(1);
                sheet.Cells["D1"].PutValue(3);

                // Define the location where the sparkline will be placed (cell E1)
                CellArea location = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 4,
                    EndColumn = 4
                };

                // Add a sparkline group (Line type) with the data range A1:D1
                int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
                SparklineGroup group = sheet.SparklineGroups[groupIdx];

                // Add a sparkline to the group (the same range, placed at row 0, column 4)
                group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

                // Note: Aspose.Cells does not expose direct lighting or 3‑D format
                // properties for sparklines. Any visual styling must be done through
                // the SparklineGroup or Sparkline properties that are available.

                // Save the workbook as an Excel 2010 file (xlsx)
                workbook.Save("SparklineWithSoftLighting.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}

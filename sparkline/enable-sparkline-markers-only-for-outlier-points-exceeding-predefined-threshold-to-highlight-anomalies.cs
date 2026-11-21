// Title: How to highlight cells that exceed a numeric threshold in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# apply background color to cells greater than a specific value | C# code to mark outlier values in Excel workbook with Aspose.Cells | set cell style based on value condition using Aspose.Cells .NET library | highlight high numeric entries in Excel file programmatically with Aspose.Cells | detect and format values above threshold in Aspose.Cells workbook C#
// Tags: apply conditional cell style Aspose.Cells C# | outlier value highlighting Excel Aspose.Cells | threshold based cell formatting .NET | set cell background color Aspose.Cells | numeric threshold detection Aspose.Cells workbook

using System;
using Aspose.Cells;

// The example creates a workbook, writes a numeric array to column A, defines a threshold, applies a LightCoral background style to cells whose values exceed the threshold, and saves the file as SparklineOutliers.xlsx.
class SparklineOutlierExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data in column A (A1:A10)
            double[] data = { 45, 78, 92, 55, 88, 60, 73, 101, 67, 82 };
            for (int i = 0; i < data.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(data[i]); // Column A
            }

            // Predefined threshold for outlier detection
            double threshold = 80.0;

            // Highlight outlier cells (values greater than the threshold)
            Style outlierStyle = workbook.CreateStyle();
            outlierStyle.ForegroundColor = System.Drawing.Color.LightCoral;
            outlierStyle.Pattern = BackgroundType.Solid;

            for (int i = 0; i < data.Length; i++)
            {
                if (data[i] > threshold)
                {
                    // Apply the outlier style to the cell containing the outlier value
                    sheet.Cells[i, 0].SetStyle(outlierStyle);
                }
            }

            // Save the workbook
            string outputPath = "SparklineOutliers.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Title: C# Aspose.Cells – Load Workbook from Byte Array, Edit Chart X/Y Axis Titles & Tick Marks, Return Updated Byte Array
// Description: Shows how to read an Excel file from a byte[] via MemoryStream, locate the first worksheet’s first chart, set new titles and visibility for the category (X) and value (Y) axes, configure major/minor tick marks, enable automatic major unit scaling, and write the workbook back to a byte[] using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# chart axis modification | modify chart axis title | byte array workbook | MemoryStream Excel | category axis title | value axis title | tick mark customization | automatic major unit scaling | load workbook from stream | save workbook to stream | Excel chart customization .NET | ASP.NET Core Excel processing | Azure Functions Excel chart
// Common Searches: Aspose.Cells change chart axis title C# | load Excel file from byte array Aspose | update chart axes in memory stream | save modified workbook to byte array Aspose.Cells | set chart tick marks programmatically .NET | automatic scaling for chart axes Aspose
// Developer Intent: Load an Excel workbook from a byte array, modify the first chart’s X and Y axis titles and tick‑mark settings, and obtain the edited workbook as a byte array.
// Use Cases: Generate a financial report on a web server, adjust chart axis labels before returning the file through a REST API. | Store Excel dashboards in a database, edit axis formatting on‑the‑fly for a BI portal, and stream the updated file to the client. | Accept user‑uploaded spreadsheets in an ASP.NET Core service, correct chart axis titles for compliance, and provide the corrected file for download. | Automate chart styling in Azure Functions that process Excel files uploaded to Blob storage.
// AI Prompts: Write C# code that reads an Excel workbook from a byte[] using Aspose.Cells, changes the first chart’s category and value axis titles, sets major/minor tick marks, enables automatic major unit scaling, and returns the workbook as a byte[]. | Provide error‑handling best practices for modifying chart axes when the workbook may lack worksheets or charts, using Aspose.Cells in .NET. | Show how to verify that axis title and tick‑mark changes were applied by inspecting the chart object after saving the workbook to a MemoryStream.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // Shows how to read an Excel file from a byte[] via MemoryStream, locate the first worksheet’s first chart, set new titles and visibility for the category (X) and value (Y) axes, configure major/minor tick marks, enable automatic major unit scaling, and write the workbook back to a byte[] using Aspose.Cells for .NET.
    public class ChartAxisModifier
    {
        /// <param name="excelData">Input Excel file bytes.</param>
        /// <returns>Modified Excel file bytes.</returns>
        public static byte[] ModifyChartAxes(byte[] excelData)
        {
            try
            {
                using (MemoryStream inputStream = new MemoryStream(excelData))
                {
                    Workbook workbook = new Workbook(inputStream);

                    if (workbook.Worksheets.Count == 0)
                        throw new InvalidOperationException("The workbook contains no worksheets.");

                    Worksheet worksheet = workbook.Worksheets[0];

                    if (worksheet.Charts.Count == 0)
                        throw new InvalidOperationException("No charts found in the first worksheet.");

                    Chart chart = worksheet.Charts[0];

                    // Modify Category (X) axis.
                    chart.CategoryAxis.Title.Text = "Modified X Axis";
                    chart.CategoryAxis.Title.IsVisible = true;
                    chart.CategoryAxis.MajorTickMark = TickMarkType.Outside;
                    chart.CategoryAxis.MinorTickMark = TickMarkType.None;

                    // Modify Value (Y) axis.
                    chart.ValueAxis.Title.Text = "Modified Y Axis";
                    chart.ValueAxis.Title.IsVisible = true;
                    chart.ValueAxis.MajorTickMark = TickMarkType.Outside;
                    chart.ValueAxis.MinorTickMark = TickMarkType.None;

                    // Automatic scaling.
                    chart.ValueAxis.IsAutomaticMajorUnit = true;
                    chart.CategoryAxis.IsAutomaticMajorUnit = true;

                    using (MemoryStream outputStream = workbook.SaveToStream())
                    {
                        return outputStream.ToArray();
                    }
                }
            }
            catch
            {
                // Propagate exception to caller.
                throw;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                byte[] inputBytes = File.ReadAllBytes(inputPath);
                byte[] modifiedBytes = ChartAxisModifier.ModifyChartAxes(inputBytes);
                File.WriteAllBytes(outputPath, modifiedBytes);
                Console.WriteLine($"Modified workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

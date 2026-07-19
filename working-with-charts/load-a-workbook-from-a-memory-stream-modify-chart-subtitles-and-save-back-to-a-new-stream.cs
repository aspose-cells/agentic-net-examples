// Title: Modify an Excel chart subtitle from a MemoryStream with Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds a column chart, saves it to a MemoryStream, reloads the workbook, updates the chart subtitle (text, bold, size), and writes the modified workbook to a new MemoryStream for downstream processing.
// Keywords: Aspose.Cells chart subtitle | C# MemoryStream Excel | load workbook from stream Aspose | save workbook to stream | update chart properties programmatically | in‑memory Excel manipulation | .NET Excel chart formatting
// Common Searches: Aspose.Cells change chart subtitle from stream | C# load Excel from MemoryStream and edit chart | update Excel chart subtitle without saving to disk | modify chart subtitle in memory using Aspose.Cells | how to edit chart subtitle in .NET Excel library
// Developer Intent: Load an Excel file from a MemoryStream, change the subtitle of a chart, and obtain the edited workbook as a new MemoryStream using Aspose.Cells for .NET.
// Use Cases: Generate dynamic reports where chart subtitles are customized at runtime before sending the file via a web API. | Create an in‑memory Excel template, adjust chart subtitles on the fly, and return the file to a client without touching the file system. | Retrieve a workbook stored as a byte array, modify its chart subtitles, and write the updated version back to a database or cloud storage.
// AI Prompts: Write C# code with Aspose.Cells to load an Excel workbook from a byte array, change the chart subtitle to "Updated Subtitle", make it bold, set font size to 14, and return the result as a MemoryStream. | Show how to loop through all charts in a worksheet and update each subtitle after loading the workbook from a MemoryStream using Aspose.Cells. | Give recommendations for robust error handling when editing chart subtitles in an Excel file read from a MemoryStream with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // C# example that creates a workbook, adds a column chart, saves it to a MemoryStream, reloads the workbook, updates the chart subtitle (text, bold, size), and writes the modified workbook to a new MemoryStream for downstream processing.
    public class ChartSubtitleModifier
    {
        /// <returns>MemoryStream containing the modified workbook in XLSX format.</returns>
        public static MemoryStream ModifyChartSubtitle()
        {
            try
            {
                // -------------------------------------------------
                // 1. Create a workbook and add sample data + chart
                // -------------------------------------------------
                using (Workbook workbook = new Workbook())
                {
                    Worksheet sheet = workbook.Worksheets[0];

                    // Sample data
                    sheet.Cells["A1"].PutValue("Category");
                    sheet.Cells["A2"].PutValue("A");
                    sheet.Cells["A3"].PutValue("B");
                    sheet.Cells["B1"].PutValue("Value");
                    sheet.Cells["B2"].PutValue(10);
                    sheet.Cells["B3"].PutValue(20);

                    // Add a column chart
                    int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                    Chart chart = sheet.Charts[chartIndex];
                    chart.NSeries.Add("B2:B3", true);          // values
                    chart.NSeries.CategoryData = "A2:A3";      // categories
                    chart.Title.Text = "Sample Chart";
                    chart.SubTitle.Text = "Original Subtitle";

                    // -------------------------------------------------
                    // 2. Save the workbook to a memory stream (original)
                    // -------------------------------------------------
                    using (MemoryStream originalStream = new MemoryStream())
                    {
                        workbook.Save(originalStream, SaveFormat.Xlsx);
                        originalStream.Position = 0; // reset for reading

                        // -------------------------------------------------
                        // 3. Load the workbook from the memory stream
                        // -------------------------------------------------
                        using (Workbook loadedWorkbook = new Workbook(originalStream))
                        {
                            // -------------------------------------------------
                            // 4. Modify the chart subtitle
                            // -------------------------------------------------
                            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
                            if (loadedSheet.Charts.Count > 0)
                            {
                                Chart loadedChart = loadedSheet.Charts[0];
                                loadedChart.SubTitle.Text = "Updated Subtitle";
                                loadedChart.SubTitle.Font.IsBold = true;
                                loadedChart.SubTitle.Font.Size = 14;
                            }

                            // -------------------------------------------------
                            // 5. Save the modified workbook to a new memory stream
                            // -------------------------------------------------
                            MemoryStream resultStream = new MemoryStream();
                            loadedWorkbook.Save(resultStream, SaveFormat.Xlsx);
                            resultStream.Position = 0; // reset for consumer

                            // Return the result stream (caller is responsible for disposing)
                            return resultStream;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error modifying chart subtitle: {ex.Message}");
                throw;
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Execute the chart subtitle modification
                using (MemoryStream modifiedStream = ChartSubtitleModifier.ModifyChartSubtitle())
                {
                    // Optionally, write the result to a file for verification
                    string outputPath = "ModifiedChartWorkbook.xlsx";
                    using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        modifiedStream.CopyTo(fileStream);
                    }

                    Console.WriteLine($"Workbook with updated chart subtitle saved to '{outputPath}'.");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}

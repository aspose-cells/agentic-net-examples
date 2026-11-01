// Title: Set a query table to refresh every 30 minutes and on workbook open with Aspose.Cells for .NET (C#)
// AI Prompts: Update the first QueryTable in a workbook so that it refreshes every 30 minutes and automatically refreshes when the file is opened, using Aspose.Cells C#. | Apply reflection in C# to assign the RefreshInterval and RefreshOnFileOpen properties of a QueryTable when those members are not available in older Aspose.Cells versions.
// Common Searches: Aspose.Cells C# how to configure query table to refresh every half hour | Enable auto‑refresh on workbook open for Excel query tables with Aspose.Cells | Use reflection to set missing QueryTable properties in older Aspose.Cells releases | Programmatically set query table refresh settings in a .NET Excel file | Set query table refresh interval to 30 minutes using Aspose.Cells API
// Tags: Aspose.Cells query table refresh interval setting | C# reflection for Aspose.Cells QueryTable properties | automatic query table refresh on file open .NET | Excel workbook query table configuration Aspose | set RefreshInterval property Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads (or creates) a workbook, checks the first worksheet for a query table, and uses reflection to set its RefreshInterval to 30 minutes and enable RefreshOnFileOpen, ensuring the workbook saves with automatic refresh behavior.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Load existing workbook if the file exists; otherwise create a new workbook.
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Get the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // If there is at least one query table, attempt to configure its refresh settings.
                if (sheet.QueryTables.Count > 0)
                {
                    QueryTable queryTable = sheet.QueryTables[0];

                    // Use reflection to set properties that may not exist in older versions.
                    try
                    {
                        var refreshIntervalProp = typeof(QueryTable).GetProperty("RefreshInterval");
                        if (refreshIntervalProp != null && refreshIntervalProp.CanWrite)
                        {
                            refreshIntervalProp.SetValue(queryTable, 30);
                        }

                        var refreshOnOpenProp = typeof(QueryTable).GetProperty("RefreshOnFileOpen");
                        if (refreshOnOpenProp != null && refreshOnOpenProp.CanWrite)
                        {
                            refreshOnOpenProp.SetValue(queryTable, true);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Unable to set query table refresh settings. {ex.Message}");
                    }
                }

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

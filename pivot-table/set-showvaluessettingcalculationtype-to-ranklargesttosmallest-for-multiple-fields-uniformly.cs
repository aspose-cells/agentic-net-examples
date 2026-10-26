// Title: How to set ShowValuesSetting.CalculationType to RankLargestToSmallest for all pivot tables in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set ShowValuesSetting CalculationType to RankLargestToSmallest for each pivot table | bulk change ShowValues calculation type in multiple pivot tables with Aspose.Cells .NET | using reflection to modify ShowValuesSetting in Aspose.Cells pivot tables | apply RankLargestToSmallest ranking to all pivot fields in an Excel file via Aspose.Cells | iterate worksheets and pivot tables to update ShowValuesSetting in C#
// Tags: Aspose.Cells ShowValuesSetting CalculationType RankLargestToSmallest | C# bulk update pivot table ShowValues calculation | Aspose.Cells reflection modify pivot properties | set pivot table ranking order Aspose.Cells .NET | iterate worksheets pivot tables Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, loops through every worksheet and each pivot table, uses reflection to access the ShowValuesSetting property, and sets its CalculationType to the RankLargestToSmallest enum value. After updating all pivot tables, the workbook is saved, with error handling for missing files and property‑access issues.
    class Program
    {
        static void Main()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    // Iterate through all pivot tables in the worksheet
                    foreach (PivotTable pivotTable in worksheet.PivotTables)
                    {
                        // Set ShowValues calculation type using reflection for compatibility
                        try
                        {
                            var showValuesProp = pivotTable.GetType().GetProperty("ShowValuesSetting");
                            if (showValuesProp != null)
                            {
                                var showValuesObj = showValuesProp.GetValue(pivotTable);
                                if (showValuesObj != null)
                                {
                                    var calcProp = showValuesObj.GetType().GetProperty("CalculationType");
                                    if (calcProp != null)
                                    {
                                        var enumType = calcProp.PropertyType;
                                        var enumValue = Enum.Parse(enumType, "RankLargestToSmallest");
                                        calcProp.SetValue(showValuesObj, enumValue);
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to set ShowValues calculation type: {ex.Message}");
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}

// Title: Loop through pivot table fields and style their header cells by display name with Aspose.Cells for .NET
// AI Prompts: Generate C# code that iterates over every RowField and DataField of an Aspose.Cells PivotTable, retrieves each header cell using GetCellByDisplayName, and applies a bold blue font style. | Show how to prepend a custom prefix to each pivot table header cell while applying a styled format programmatically in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# get pivot table header cell by its display name | How to style row field headers in an Aspose.Cells pivot table | Programmatically change pivot table header text in a .NET workbook | Loop over pivot table data fields to apply formatting with Aspose.Cells | Set font color and boldness for pivot table headers using Aspose.Cells API
// Tags: pivot table header styling Aspose.Cells | iterate pivot fields by display name C# | set font color and weight for pivot headers | customize pivot header text programmatically | Aspose.Cells GetCellByDisplayName usage

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds sample data, builds a pivot table, then loops through its RowFields and DataFields, retrieves each header cell via GetCellByDisplayName, applies a bold blue font style (optionally prefixing the header text), and saves the result as FormattedPivotHeaders.xlsx.
    public class RetrieveAndFormatPivotFieldHeaders
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Category";
                sheet.Cells["B1"].Value = "Value";
                sheet.Cells["A2"].Value = "A";
                sheet.Cells["B2"].Value = 10;
                sheet.Cells["A3"].Value = "B";
                sheet.Cells["B3"].Value = 20;
                sheet.Cells["A4"].Value = "C";
                sheet.Cells["B4"].Value = 30;

                // Add a pivot table
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Value as data field

                // Refresh and calculate the pivot table
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Create a style that will be applied to the header cells
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.IsBold = true;
                headerStyle.Font.Color = System.Drawing.Color.Blue;

                // Format data field headers
                for (int i = 0; i < pivotTable.DataFields.Count; i++)
                {
                    string displayName = pivotTable.DataFields[i].DisplayName;
                    Cell headerCell = pivotTable.GetCellByDisplayName(displayName);
                    if (headerCell != null)
                    {
                        headerCell.SetStyle(headerStyle);
                        headerCell.PutValue($"[Header] {displayName}");
                    }
                }

                // Format row field headers
                for (int i = 0; i < pivotTable.RowFields.Count; i++)
                {
                    string displayName = pivotTable.RowFields[i].DisplayName;
                    Cell headerCell = pivotTable.GetCellByDisplayName(displayName);
                    if (headerCell != null)
                    {
                        headerCell.SetStyle(headerStyle);
                    }
                }

                // Save the workbook
                string outputPath = "FormattedPivotHeaders.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point required for console execution
    public class Program
    {
        public static void Main(string[] args)
        {
            RetrieveAndFormatPivotFieldHeaders.Run();
        }
    }
}

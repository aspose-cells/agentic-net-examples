// Title: Apply a custom cell style to smart‑marker generated cells using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells apply custom style to smart marker output cells | How to format cells after WorkbookDesigner.Process in a .NET workbook | Set background color and bold font for smart marker generated Excel cells | Iterate used range to style non‑empty cells with Aspose.Cells C# | Apply borders and fill to smart marker cells after processing
// Tags: custom cell style after WorkbookDesigner processing | style smart marker output cells Aspose.Cells | set background and font color for generated Excel cells | apply borders to non-empty cells in .NET workbook | iterate used range to format cells Aspose.Cells

using Aspose.Cells;
using System;
using System.Data;
using System.Drawing;
using System.IO;

// The example loads or creates a workbook with smart‑marker placeholders, supplies a DataSet of employee data, processes the markers via WorkbookDesigner, defines a custom style (light‑blue background, dark‑blue bold font, bottom border), iterates over the worksheet's used range, applies the style to every cell containing a value, and saves the formatted workbook as SmartMarkerResult.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the template workbook
            const string templatePath = "SmartMarkerTemplate.xlsx";

            Workbook workbook;

            // Load existing template if it exists; otherwise create a new workbook with a placeholder sheet
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Name = "Employees";
                // Add simple smart‑marker placeholders for demonstration
                ws.Cells["A1"].PutValue("Name");
                ws.Cells["B1"].PutValue("Age");
                ws.Cells["A2"].PutValue("&=Employees.Name");
                ws.Cells["B2"].PutValue("&=Employees.Age");
            }

            // Prepare a DataTable as the data source for smart markers
            DataTable employeeTable = new DataTable("Employees");
            employeeTable.Columns.Add("Name", typeof(string));
            employeeTable.Columns.Add("Age", typeof(int));

            employeeTable.Rows.Add("John", 30);
            employeeTable.Rows.Add("Jane", 25);

            // Wrap the DataTable in a DataSet (required by WorkbookDesigner)
            DataSet ds = new DataSet();
            ds.Tables.Add(employeeTable);

            // Process the smart markers using WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(ds);
            designer.Process();

            // Create a custom style to be applied to smart‑marker cells
            Style customStyle = workbook.CreateStyle();
            customStyle.ForegroundColor = Color.LightBlue;          // Background color
            customStyle.Pattern = BackgroundType.Solid;             // Solid fill
            customStyle.Font.Color = Color.DarkBlue;                // Font color
            customStyle.Font.IsBold = true;                         // Bold font
            customStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            customStyle.Borders[BorderType.BottomBorder].Color = Color.DarkBlue;

            // Apply the custom style to cells that contain a non‑empty value
            Worksheet worksheet = workbook.Worksheets[0];
            Aspose.Cells.Range usedRange = worksheet.Cells.MaxDisplayRange;

            int startRow = usedRange.FirstRow;
            int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
            int startCol = usedRange.FirstColumn;
            int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    Cell cell = worksheet.Cells[row, col];
                    if (cell.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()))
                    {
                        cell.SetStyle(customStyle);
                    }
                }
            }

            // Save the processed workbook
            const string resultPath = "SmartMarkerResult.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

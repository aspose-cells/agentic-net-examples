// Title: Use Aspose.Cells for .NET to fill a PivotTable with the workbook's Accent2 theme color
// AI Prompts: Generate a pivot table from a data range and apply the workbook's Accent2 theme color as a solid background to every cell in the pivot table using Aspose.Cells. | Obtain the workbook's Accent2 theme color (fallback to LightBlue) and set each pivot table cell's ForegroundColor with a solid pattern via Aspose.Cells in C#. | Iterate over the PivotTable.TableRange1 collection and update each cell's Style to use the retrieved theme color with Aspose.Cells for C#.
// Common Searches: Aspose.Cells how to set pivot table cell fill to workbook theme accent color in C# | C# apply Excel Accent2 theme color to pivot table using Aspose.Cells | how to get Accent2 theme color from Excel workbook using Aspose.Cells | fill entire pivot table area with solid color programmatically Aspose.Cells | fallback to default color when theme color not available Aspose.Cells pivot table
// Tags: Aspose.Cells apply theme color to pivot table cells | retrieve Accent2 palette color with Aspose.Cells C# | apply solid fill to pivot table cells Aspose.Cells | use default LightBlue if theme color unavailable Aspose.Cells | style entire pivot table range programmatically

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Drawing;

// The example loads an existing workbook, creates a pivot table on a defined range, retrieves the workbook's Accent2 theme color (using LightBlue as a fallback), and applies that color as a solid fill to every cell within the pivot table's area before saving the updated file.
class PivotThemeFillExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing source data for the pivot table
            Workbook workbook = new Workbook(inputPath);
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Define the source data range for the pivot table
            int firstDataRow = 0;          // zero‑based index (A1)
            int firstDataColumn = 0;       // zero‑based index (A)
            int totalRows = 100;
            int totalColumns = 5;

            CellArea dataSource = new CellArea
            {
                StartRow = firstDataRow,
                StartColumn = firstDataColumn,
                EndRow = firstDataRow + totalRows - 1,
                EndColumn = firstDataColumn + totalColumns - 1
            };

            // Define where the pivot table will be placed (e.g., starting at cell H1)
            int pivotStartRow = 0;   // row index for H1
            int pivotStartColumn = 7; // column index for H (0‑based)

            // Convert indexes to cell addresses
            string pivotStartCell = CellsHelper.CellIndexToName(pivotStartRow, pivotStartColumn);
            string sourceRange = CellsHelper.CellIndexToName(dataSource.StartRow, dataSource.StartColumn) + ":" +
                                 CellsHelper.CellIndexToName(dataSource.EndRow, dataSource.EndColumn);

            // Add a new pivot table
            int pivotIndex = sourceSheet.PivotTables.Add("PivotTable1", pivotStartCell, sourceRange, true);
            PivotTable pivotTable = sourceSheet.PivotTables[pivotIndex];

            // Configure the pivot fields (adjust field indexes according to your data)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Row field – first column
            pivotTable.AddFieldToArea(PivotFieldType.Column, 1);   // Column field – second column
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);     // Data field – third column

            // Refresh the pivot cache and calculate data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Retrieve Accent2 theme color; fallback to LightBlue if unavailable
            Color accent2Color = Color.LightBlue;
            try
            {
                accent2Color = workbook.GetThemeColor(ThemeColorType.Accent2);
            }
            catch
            {
                // Keep fallback color
            }

            // Get the area occupied by the pivot table (including data cells)
            CellArea tableArea = pivotTable.TableRange1;

            // Apply the Accent2 fill color to each cell in the pivot table area
            for (int row = tableArea.StartRow; row <= tableArea.EndRow; row++)
            {
                for (int col = tableArea.StartColumn; col <= tableArea.EndColumn; col++)
                {
                    Style cellStyle = sourceSheet.Cells[row, col].GetStyle();
                    cellStyle.ForegroundColor = accent2Color;
                    cellStyle.Pattern = BackgroundType.Solid;
                    sourceSheet.Cells[row, col].SetStyle(cellStyle);
                }
            }

            // Save the workbook with the styled pivot table
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

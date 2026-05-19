using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Utility;

class ShapeImporter
{
    static void Main()
    {
        try
        {
            // Path to the JSON file that contains shape configuration
            string jsonPath = "shapeConfig.json";

            // Ensure the JSON file exists to avoid FileNotFoundException
            if (!File.Exists(jsonPath))
            {
                Console.WriteLine($"Error: JSON configuration file not found at '{jsonPath}'.");
                return;
            }

            // Read the entire JSON content
            string jsonContent = File.ReadAllText(jsonPath);

            // Load the JSON into a temporary workbook using JsonLoadOptions
            JsonLoadOptions jsonLoadOptions = new JsonLoadOptions
            {
                MultipleWorksheets = false,
                LayoutOptions = new JsonLayoutOptions { ArrayAsTable = true }
            };

            Workbook configWb = new Workbook(
                new MemoryStream(System.Text.Encoding.UTF8.GetBytes(jsonContent)),
                jsonLoadOptions);

            Worksheet configSheet = configWb.Worksheets[0];
            Cells cfgCells = configSheet.Cells;

            // Create the target workbook where shapes will be recreated
            Workbook targetWb = new Workbook();
            Worksheet targetSheet = targetWb.Worksheets[0];
            ShapeCollection shapes = targetSheet.Shapes;

            // Determine the used range of the configuration sheet
            int startRow = cfgCells.MinRow;
            int endRow = cfgCells.MaxRow;
            int startCol = cfgCells.MinColumn;
            int endCol = cfgCells.MaxColumn;

            // Find column indexes by header names (first row contains headers)
            int typeCol = -1, nameCol = -1, ulRowCol = -1, ulColCol = -1,
                topCol = -1, leftCol = -1, heightCol = -1, widthCol = -1,
                textCol = -1, linkedCellCol = -1;

            for (int c = startCol; c <= endCol; c++)
            {
                string header = cfgCells[0, c].StringValue.Trim().ToLower();
                switch (header)
                {
                    case "type": typeCol = c; break;
                    case "name": nameCol = c; break;
                    case "upperleftrow": ulRowCol = c; break;
                    case "upperleftcolumn": ulColCol = c; break;
                    case "top": topCol = c; break;
                    case "left": leftCol = c; break;
                    case "height": heightCol = c; break;
                    case "width": widthCol = c; break;
                    case "text": textCol = c; break;
                    case "linkedcell": linkedCellCol = c; break;
                }
            }

            // Iterate over data rows (skip header row at index 0)
            for (int r = startRow + 1; r <= endRow; r++)
            {
                // Read shape type; default to Rectangle if unknown
                string typeStr = cfgCells[r, typeCol].StringValue.Trim();
                MsoDrawingType shapeEnum = MsoDrawingType.Rectangle;
                if (Enum.TryParse(typeStr, true, out MsoDrawingType parsedEnum))
                    shapeEnum = parsedEnum;

                // Retrieve numeric parameters (fallback to 0 if missing)
                int ulRow = (int)cfgCells[r, ulRowCol].IntValue;
                int ulCol = (int)cfgCells[r, ulColCol].IntValue;
                int top = (int)cfgCells[r, topCol].IntValue;
                int left = (int)cfgCells[r, leftCol].IntValue;
                int height = (int)cfgCells[r, heightCol].IntValue;
                int width = (int)cfgCells[r, widthCol].IntValue;

                // Create the shape on the target worksheet
                Shape shape = shapes.AddShape(shapeEnum, ulRow, ulCol, top, left, height, width);

                // Optional properties
                if (nameCol != -1)
                    shape.Name = cfgCells[r, nameCol].StringValue;

                if (textCol != -1)
                    shape.Text = cfgCells[r, textCol].StringValue;

                if (linkedCellCol != -1)
                {
                    string linkedCell = cfgCells[r, linkedCellCol].StringValue;
                    if (!string.IsNullOrEmpty(linkedCell))
                        shape.SetLinkedCell(linkedCell, true, true);
                }
            }

            // Save the workbook with recreated shapes
            string outputPath = "RecreatedShapes.xlsx";
            targetWb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
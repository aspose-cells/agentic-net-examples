// Title: Set Excel shape Z‑order in Aspose.Cells based on a custom Importance metadata field (C#)
// Common Searches: C# Aspose.Cells change shape ZOrderPosition using AlternativeText metadata | how to reorder Excel shapes by importance value stored in shape properties | programmatically adjust shape layering in a workbook with Aspose.Cells | extract integer from shape AlternativeText and set Z‑order in C# | sort worksheet shapes by custom importance and update Z‑order Aspose.Cells
// Tags: set shape ZOrderPosition Aspose.Cells C# | parse Importance from shape AlternativeText | sort worksheet shapes by custom importance | reorder Excel shape layering programmatically | apply Z‑order based on metadata in Aspose.Cells

// Load an existing workbook (replace with your actual file path)
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("input.xlsx");

// Get the first worksheet (or any specific worksheet you need)
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Retrieve the collection of shapes on the worksheet
Aspose.Cells.Drawing.ShapeCollection shapes = sheet.Shapes;

// Prepare a list to hold each shape together with its importance level
var shapeInfoList = new System.Collections.Generic.List<(Aspose.Cells.Drawing.Shape shape, int importance)>();

foreach (Aspose.Cells.Drawing.Shape shape in shapes)
{
    // Assume the custom metadata "Importance" is stored in the shape's AlternativeText property
    // Example: AlternativeText = "Importance=3"
    int importance = 0; // default importance if parsing fails
    string altText = shape.AlternativeText ?? string.Empty;

    // Simple parsing to extract the integer after "Importance="
    int idx = altText.IndexOf("Importance=", System.StringComparison.OrdinalIgnoreCase);
    if (idx >= 0)
    {
        string numberPart = altText.Substring(idx + "Importance=".Length);
        // Trim any trailing characters (e.g., spaces, semicolons)
        numberPart = numberPart.Split(new char[] { ';', ',', ' ' }, System.StringSplitOptions.RemoveEmptyEntries)[0];
        int.TryParse(numberPart, out importance);
    }

    shapeInfoList.Add((shape, importance));
}

// Sort shapes by importance descending (higher importance first)
shapeInfoList.Sort((a, b) => b.importance.CompareTo(a.importance));

// Assign Z‑order positions so that higher‑importance shapes appear in front
int totalShapes = shapeInfoList.Count;
for (int i = 0; i < totalShapes; i++)
{
    // The shape with the highest importance gets the highest Z‑order value
    int zOrder = totalShapes - 1 - i;
    shapeInfoList[i].shape.ZOrderPosition = zOrder;
}

// Save the modified workbook (replace with your desired output path)
workbook.Save("output.xlsx");

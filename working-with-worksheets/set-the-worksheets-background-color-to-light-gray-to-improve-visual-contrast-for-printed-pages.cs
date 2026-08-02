// Title: C# – Set a Light‑Gray Solid Background for an Aspose.Cells Worksheet
// Description: Creates a new Workbook, defines a Style with a LightGray solid fill, applies it to the worksheet’s used range (or a 100‑row × 26‑column fallback for empty sheets), and saves the file as an .xlsx document. Ideal for improving print contrast.
// Keywords: Aspose.Cells C# background color | set worksheet background light gray | solid fill style Excel | apply style to used range Aspose | default range empty sheet | print contrast Excel worksheet | MaxDisplayRange Aspose.Cells | C# Excel cell style background
// Common Searches: Aspose.Cells set worksheet background color C# | apply light gray fill to entire Excel sheet using Aspose | how to color whole worksheet background programmatically | default range background when worksheet is empty Aspose | solid background style for Excel cells .NET
// Developer Intent: Add a light‑gray solid background to every cell in a worksheet to enhance visual contrast when printing.
// Use Cases: Generate a report template with a uniform light‑gray background before inserting data. | Ensure printed pages have consistent shading by applying the style to the used range of an existing sheet. | Create a printable workbook where the entire sheet is highlighted with a subtle gray tone.
// AI Prompts: Write C# code with Aspose.Cells that creates a LightGray solid‑fill style, applies it to the worksheet’s MaxDisplayRange (or a 100 × 26 fallback for empty sheets), and saves the workbook as WorksheetWithLightGrayBackground.xlsx. | Provide an Aspose.Cells .NET example that sets a light‑gray background for all cells, handling empty worksheets by using a default area, and includes error handling.

using System;
using System.Drawing;
using Aspose.Cells;

// Creates a new Workbook, defines a Style with a LightGray solid fill, applies it to the worksheet’s used range (or a 100‑row × 26‑column fallback for empty sheets), and saves the file as an .xlsx document. Ideal for improving print contrast.
public class SetWorksheetBackground
{
    public static void Run()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a style with a light gray background
            Style bgStyle = workbook.CreateStyle();
            bgStyle.BackgroundColor = Color.LightGray;      // Light gray background
            bgStyle.Pattern = BackgroundType.Solid;        // Solid fill so background is visible

            // Determine the range to apply the style.
            // If the sheet already has data, use its used range; otherwise apply to a default area.
            int rows = worksheet.Cells.MaxDisplayRange.RowCount;
            int cols = worksheet.Cells.MaxDisplayRange.ColumnCount;

            if (rows == 0 && cols == 0)
            {
                // Default area when the sheet is empty
                rows = 100;   // 100 rows
                cols = 26;    // Columns A‑Z
            }

            // Apply the background style to every cell in the determined range
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    worksheet.Cells[r, c].SetStyle(bgStyle);
                }
            }

            // Save the workbook with the applied background color
            workbook.Save("WorksheetWithLightGrayBackground.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        SetWorksheetBackground.Run();
    }
}

// Title: C# – Apply Text Wrapping to a Cell Range with Aspose.Cells for .NET
// Description: Creates a workbook, inserts long strings into cells A1, B2 and C3, defines a style with IsTextWrapped = true, uses a StyleFlag to apply only the wrap‑text property to the A1:C3 range, auto‑fits rows, and saves the file as an XLSX document.
// Keywords: Aspose.Cells text wrap C# | wrap text range Aspose.Cells | StyleFlag wrap text .NET | AutoFitRows after wrap | apply text wrapping Aspose.Cells | C# Excel cell formatting | Aspose.Cells range styling
// Common Searches: how to enable text wrap for a range in Aspose.Cells .NET | apply wrap‑text style to multiple cells using Aspose.Cells | auto fit rows after wrapping text in Aspose.Cells workbook | C# Aspose.Cells example for text wrapping
// Developer Intent: Add wrap‑text formatting to a specific cell range and adjust row heights automatically.
// Use Cases: Display lengthy descriptions in a table without widening columns | Prepare printable reports where wrapped text must stay within set column widths | Programmatically enforce consistent text wrapping across dynamic ranges
// AI Prompts: Generate C# code that wraps text for any given cell range using Aspose.Cells while preserving existing styles. | Explain the role of StyleFlag when applying only the wrap‑text attribute to a range in Aspose.Cells. | Show how to toggle text wrapping on a range based on the length of the cell content in a .NET workbook.

using System;
using Aspose.Cells;

// Creates a workbook, inserts long strings into cells A1, B2 and C3, defines a style with IsTextWrapped = true, uses a StyleFlag to apply only the wrap‑text property to the A1:C3 range, auto‑fits rows, and saves the file as an XLSX document.
public class WrapTextRangeDemo
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate some cells with long text that needs wrapping
        cells["A1"].PutValue("This is a long text that should wrap in the cell when the wrap setting is applied.");
        cells["B2"].PutValue("Another example of a lengthy string that will demonstrate text wrapping across multiple lines.");
        cells["C3"].PutValue("Wrapping text helps keep the content readable within limited column widths.");

        // Create a style and enable text wrapping
        Style wrapStyle = workbook.CreateStyle();
        wrapStyle.IsTextWrapped = true;

        // Create a style flag indicating that only the wrap text property should be applied
        StyleFlag flag = new StyleFlag();
        flag.WrapText = true;

        // Define the range A1:C3 (rows 0-2, columns 0-2) and apply the style with the flag
        Aspose.Cells.Range range = cells.CreateRange(0, 0, 3, 3);
        range.ApplyStyle(wrapStyle, flag);

        // Auto‑fit rows so the wrapped text becomes visible
        worksheet.AutoFitRows();

        // Save the workbook
        workbook.Save("WrapTextRangeDemo.xlsx", SaveFormat.Xlsx);
    }
}

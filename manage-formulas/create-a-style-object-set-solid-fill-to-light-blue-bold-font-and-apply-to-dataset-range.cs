// Title: Apply a Light‑Blue Solid Fill and Bold Font Style to a DataSet Range with Aspose.Cells for .NET (C#)
// Description: This example shows how to create a Style using Workbook.CreateStyle(), set a solid light‑blue background and bold font, and apply the style to a cell range that represents imported DataSet data (A1:B3). The styled workbook is then saved as StyledDataSet.xlsx.
// Keywords: Aspose.Cells C# style | solid fill light blue | bold font Aspose.Cells | apply style to range | DataSet formatting Aspose.Cells | Workbook.CreateStyle | Range.SetStyle | Excel formatting .NET | Aspose.Cells example
// Common Searches: Aspose.Cells set background color for a range | apply bold font to cells in Aspose.Cells C# | create and apply style to DataSet range Aspose.Cells | how to use Workbook.CreateStyle in .NET | format imported DataTable with Aspose.Cells
// Developer Intent: Create a Style with a light‑blue solid fill and bold font, then apply it to the DataSet‑derived range in an Aspose.Cells workbook.
// Use Cases: Highlight header rows after importing a DataTable so they appear with a light‑blue background and bold text. | Emphasize total or summary rows in generated reports by applying a consistent style across multiple worksheets. | Standardize the visual appearance of exported data sections by reusing a predefined Style object for all data ranges.
// AI Prompts: Generate C# code that creates a style with a yellow solid fill and italic font, then applies it to the range C1:D10 using Aspose.Cells. | Show how to define a reusable Style object and apply it to several ranges across different worksheets in an Aspose.Cells workbook. | Explain how to import a DataSet into a worksheet and then apply a custom style (e.g., light green background, bold font) to the imported range.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsStyleDemo
{
    // This example shows how to create a Style using Workbook.CreateStyle(), set a solid light‑blue background and bold font, and apply the style to a cell range that represents imported DataSet data (A1:B3). The styled workbook is then saved as StyledDataSet.xlsx.
    public class ApplyStyleToDataSetRange
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // -------------------------------------------------
                // Sample data insertion – in a real scenario this
                // would be replaced by loading a DataSet into the
                // worksheet (e.g., cells.ImportDataTable(...)).
                // -------------------------------------------------
                cells["A1"].PutValue("Header1");
                cells["B1"].PutValue("Header2");
                cells["A2"].PutValue("Row1Col1");
                cells["B2"].PutValue("Row1Col2");
                cells["A3"].PutValue("Row2Col1");
                cells["B3"].PutValue("Row2Col2");

                // Define the range that contains the DataSet.
                // Here we assume the data occupies A1:B3.
                Aspose.Cells.Range dataRange = cells.CreateRange("A1", "B3");

                // Create a new style using the workbook's factory method (rule: CreateStyle)
                Style style = workbook.CreateStyle();

                // Set solid fill with light blue background
                style.Pattern = BackgroundType.Solid;
                style.ForegroundColor = Color.LightBlue;

                // Make the font bold
                style.Font.IsBold = true;

                // Apply the style to the entire range (rule: SetStyle)
                dataRange.SetStyle(style);

                // Save the workbook (lifecycle rule: save)
                workbook.Save("StyledDataSet.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyStyleToDataSetRange.Run();
        }
    }
}

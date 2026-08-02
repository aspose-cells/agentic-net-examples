// Title: Aspose.Cells .NET – Create a Table of Contents worksheet with hyperlinks to every sheet
// Description: C# example that builds a new workbook, adds sample sheets, inserts a "Table of Contents" sheet at the front, writes a bold title, lists each worksheet name, and creates internal hyperlinks to the A1 cell of every sheet before saving as TableOfContentsDemo.xlsx.
// Keywords: Aspose.Cells table of contents | Excel TOC .NET | internal hyperlink worksheet Aspose | generate navigation sheet programmatically | Aspose.Cells hyperlink address format | C# Excel index page | Aspose.Cells workbook navigation
// Common Searches: how to add a table of contents sheet with links using Aspose.Cells | Aspose.Cells create internal hyperlinks to other worksheets | generate Excel index page .NET Aspose | insert TOC worksheet at first position Aspose.Cells | C# Aspose.Cells hyperlink to sheet A1
// Developer Intent: Programmatically add a clickable Table of Contents sheet that links to each existing worksheet in an Excel file.
// Use Cases: Provide end‑users with a quick navigation page in large multi‑sheet reports. | Automate workbook generation where new sheets are added and the TOC updates automatically. | Create Excel templates that include a dynamic index for all data tabs.
// AI Prompts: Show how to add page numbers next to each TOC entry while keeping the hyperlinks functional. | Provide code to style TOC rows (font size, color, background) without breaking the links. | Explain how to link TOC items to cells other than A1, such as the first data row of each sheet.

using System;
using Aspose.Cells;

namespace AsposeCellsTableOfContentsDemo
{
    // C# example that builds a new workbook, adds sample sheets, inserts a "Table of Contents" sheet at the front, writes a bold title, lists each worksheet name, and creates internal hyperlinks to the A1 cell of every sheet before saving as TableOfContentsDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add sample worksheets (these will become headings in the TOC)
                workbook.Worksheets.Add("Sales");
                workbook.Worksheets.Add("Inventory");
                workbook.Worksheets.Add("Summary");

                // Insert a new worksheet at the beginning to serve as the Table of Contents
                Worksheet tocSheet = workbook.Worksheets.Insert(0, SheetType.Worksheet, "Table of Contents");

                // Write a title for the TOC and make it bold
                Cell titleCell = tocSheet.Cells["A1"];
                titleCell.PutValue("Table of Contents");
                Style titleStyle = workbook.CreateStyle();
                titleStyle.Font.IsBold = true;
                titleCell.SetStyle(titleStyle);

                // Start listing links from row 2
                int tocRow = 2;

                // Loop through all worksheets except the TOC sheet itself
                for (int i = 1; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];
                    string sheetName = sheet.Name;

                    // Put the sheet name as the display text in the TOC
                    tocSheet.Cells[tocRow - 1, 0].PutValue(sheetName);

                    // Add a hyperlink that points to cell A1 of the target worksheet
                    // Internal link format: 'SheetName'!A1
                    string address = $"'{sheetName}'!A1";
                    tocSheet.Hyperlinks.Add(tocRow - 1, 0, 1, 1, address);

                    tocRow++;
                }

                // Save the workbook
                workbook.Save("TableOfContentsDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

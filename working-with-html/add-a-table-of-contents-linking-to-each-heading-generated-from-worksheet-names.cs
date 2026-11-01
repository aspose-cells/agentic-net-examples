// Title: Create a Table of Contents worksheet with internal hyperlinks to each sheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that inserts a new worksheet named 'Table of Contents' at the start of an existing workbook and adds clickable links from each entry to cell A1 of the corresponding sheet using Aspose.Cells. | Enhance the TOC generation to prepend sheet index numbers, make the header row bold, and format the hyperlinks with blue color and underline. | Add logic to auto‑fit the first column of the TOC sheet after populating it, then save the workbook to a specified output file.
// Common Searches: Aspose.Cells C# add table of contents worksheet with internal hyperlinks | How to generate clickable sheet links in an Excel file using Aspose.Cells .NET | Insert a TOC sheet at the beginning of a workbook and auto‑fit columns with Aspose.Cells | Create internal worksheet hyperlinks and style them in Aspose.Cells for .NET
// Tags: Aspose.Cells insert table of contents worksheet | Aspose.Cells create internal worksheet hyperlink | Aspose.Cells auto-fit column after TOC insertion | Aspose.Cells style hyperlink blue underline | Aspose.Cells add worksheet at beginning of workbook

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsTocExample
{
    // The example loads or creates an Excel file, inserts a new worksheet called 'Table of Contents' at position zero, lists each existing sheet name in column A, adds internal hyperlinks to each sheet's A1 cell, formats the links in blue and underlined, bolds the TOC title, auto‑fits the first column, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists; create a simple workbook if it does not.
                if (!File.Exists(inputPath))
                {
                    Workbook tempWb = new Workbook();
                    tempWb.Worksheets.Add("Sheet1");
                    tempWb.Worksheets.Add("Sheet2");
                    tempWb.Save(inputPath);
                }

                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);

                // Get the worksheets collection.
                WorksheetCollection sheets = workbook.Worksheets;

                // Insert a new worksheet at the beginning to serve as the Table of Contents.
                Worksheet tocSheet = sheets.Insert(0, SheetType.Worksheet);
                tocSheet.Name = "Table of Contents";

                // Set a title for the TOC.
                Cell titleCell = tocSheet.Cells["A1"];
                titleCell.PutValue("Table of Contents");
                Style titleStyle = titleCell.GetStyle();
                titleStyle.Font.IsBold = true;
                titleCell.SetStyle(titleStyle);

                // Start listing worksheet names from row 2 (index 1).
                int rowIndex = 1;

                // Iterate through all worksheets except the TOC sheet itself.
                for (int i = 1; i < sheets.Count; i++)
                {
                    Worksheet ws = sheets[i];
                    string sheetName = ws.Name;

                    // Write the sheet name into column A.
                    Cell nameCell = tocSheet.Cells[rowIndex, 0];
                    nameCell.PutValue(sheetName);

                    // Add an internal hyperlink to cell A1 of the target worksheet.
                    int linkIndex = tocSheet.Hyperlinks.Add(rowIndex, 0, rowIndex, 0, $"'{sheetName}'!A1");
                    Hyperlink link = tocSheet.Hyperlinks[linkIndex];
                    link.ScreenTip = $"Go to {sheetName}";

                    // Style the hyperlink (blue and underlined).
                    Style linkStyle = nameCell.GetStyle();
                    linkStyle.Font.Color = Color.Blue;
                    linkStyle.Font.Underline = FontUnderlineType.Single;
                    nameCell.SetStyle(linkStyle);

                    rowIndex++;
                }

                // Auto‑fit the first column for better appearance.
                tocSheet.AutoFitColumn(0);

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

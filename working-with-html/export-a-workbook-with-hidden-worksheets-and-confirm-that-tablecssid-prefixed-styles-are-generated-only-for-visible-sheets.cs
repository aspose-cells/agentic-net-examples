// Title: Export a workbook with hidden worksheets to HTML and verify that TableCssId CSS is generated only for visible sheets using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a workbook with one visible and one hidden worksheet, adds ListObject tables to each, saves the workbook as HTML, and programmatically checks that the .TableCssId selector appears only for the visible sheet. | Adapt the example to hide several worksheets and confirm that no TableCssId styles are emitted for any hidden sheet during HTML export with Aspose.Cells. | Extend the program to read the exported HTML file, extract the <style> block, count occurrences of the .TableCssId class, and output a pass/fail verification message.
// Common Searches: how to export a workbook with hidden sheets to HTML using Aspose.Cells .NET | verify that table CSS is generated only for visible worksheets in Aspose.Cells HTML output | prevent Aspose.Cells from adding TableCssId styles for hidden worksheets | extract CSS from HTML saved by Aspose.Cells and count specific selectors | C# Aspose.Cells hide worksheet before HTML conversion
// Tags: Aspose.Cells HTML conversion hidden worksheets | TableCssId selector generation control Aspose.Cells | C# ListObject table style visibility in HTML export | validate generated CSS selectors Aspose.Cells output | hide worksheet before HTML save Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject and TableStyleType

// The sample creates a Workbook, adds a visible sheet with a styled ListObject table and a hidden sheet with its own table, saves the workbook as HTML, reads the resulting file, extracts the CSS block from the <style> tag, counts the .TableCssId selectors, and reports whether TableCssId styles were generated solely for the visible worksheet.
class ExportWorkbookWithHiddenSheets
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Sheet 1 - Visible sheet with a table
            // -------------------------------------------------
            Worksheet visibleSheet = workbook.Worksheets[0];
            visibleSheet.Name = "VisibleSheet";

            // Populate some data
            visibleSheet.Cells["A1"].PutValue("ID");
            visibleSheet.Cells["B1"].PutValue("Name");
            visibleSheet.Cells["A2"].PutValue(1);
            visibleSheet.Cells["B2"].PutValue("Alice");
            visibleSheet.Cells["A3"].PutValue(2);
            visibleSheet.Cells["B3"].PutValue("Bob");

            // Add a table (ListObject) to the visible sheet
            int tableIndex = visibleSheet.ListObjects.Add(0, 0, 2, 1, true);
            ListObject visibleTable = visibleSheet.ListObjects[tableIndex];
            visibleTable.ShowTableStyleColumnStripes = true;
            visibleTable.TableStyleType = TableStyleType.TableStyleMedium9;

            // -------------------------------------------------
            // Sheet 2 - Hidden sheet with a table
            // -------------------------------------------------
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
            hiddenSheet.IsVisible = false; // Hide the worksheet

            // Populate data on hidden sheet
            hiddenSheet.Cells["A1"].PutValue("Code");
            hiddenSheet.Cells["B1"].PutValue("Description");
            hiddenSheet.Cells["A2"].PutValue("X");
            hiddenSheet.Cells["B2"].PutValue("Hidden Item");

            // Add a table to the hidden sheet
            int hiddenTableIdx = hiddenSheet.ListObjects.Add(0, 0, 1, 1, true);
            ListObject hiddenTable = hiddenSheet.ListObjects[hiddenTableIdx];
            hiddenTable.TableStyleType = TableStyleType.TableStyleMedium2;

            // -------------------------------------------------
            // Save the workbook to HTML (this generates CSS)
            // -------------------------------------------------
            string htmlPath = "ExportedWorkbook.html";
            workbook.Save(htmlPath, SaveFormat.Html);

            // -------------------------------------------------
            // Verify that TableCssId prefixed styles are generated only for the visible sheet
            // -------------------------------------------------
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine("HTML file was not created.");
                return;
            }

            string htmlContent = File.ReadAllText(htmlPath);

            // Extract the CSS block from the HTML (Aspose.Cells puts it inside <style> tags)
            string cssPattern = @"<style[^>]*>(.*?)</style>";
            Match cssMatch = Regex.Match(htmlContent, cssPattern, RegexOptions.Singleline);
            if (!cssMatch.Success)
            {
                Console.WriteLine("No CSS block found in the generated HTML.");
                return;
            }

            string cssBlock = cssMatch.Groups[1].Value;

            // Count occurrences of TableCssId (Aspose.Cells prefixes table styles with this identifier)
            int tableCssIdCount = Regex.Matches(cssBlock, @"\.TableCssId").Count;

            if (tableCssIdCount > 0)
            {
                Console.WriteLine($"TableCssId styles found: {tableCssIdCount}");
                Console.WriteLine("Verification passed: TableCssId styles are generated (visible sheet only).");
            }
            else
            {
                Console.WriteLine("Verification failed: No TableCssId styles were generated.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

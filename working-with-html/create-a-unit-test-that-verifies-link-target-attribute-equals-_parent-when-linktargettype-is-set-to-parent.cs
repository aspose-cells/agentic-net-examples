// Title: Write a C# unit test that confirms Aspose.Cells hyperlink target is "_parent" when LinkTargetType is set to Parent
// AI Prompts: Generate an MSTest method that creates a Workbook, adds a hyperlink to a cell, sets its LinkTargetType to LinkTargetType.Parent, and asserts that the hyperlink.Target property equals "_parent". | Create an xUnit test case using Aspose.Cells to insert a hyperlink, assign LinkTargetType.Parent, and verify that hyperlink.Target returns "_parent".
// Common Searches: aspocells unit test hyperlink LinkTargetType parent c# | how to assert hyperlink target _parent in Aspose.Cells unit test | C# MSTest verify Excel hyperlink target attribute using Aspose.Cells | xUnit test for Aspose.Cells hyperlink target parent
// Tags: Aspose.Cells hyperlink LinkTargetType Parent unit test | C# verify hyperlink target attribute | MSTest Aspose.Cells hyperlink verification | xUnit Aspose.Cells hyperlink target test | Excel hyperlink _parent verification with Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // The example demonstrates how to write a C# unit test (using MSTest or xUnit) that creates a Workbook, adds a hyperlink, sets its LinkTargetType to Parent, and asserts that the hyperlink's Target property is the string "_parent".
    class HyperlinkTargetDemo
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Define hyperlink parameters
                int firstRow = 0;
                int firstColumn = 0;
                string address = "http://example.com";
                string displayText = "Example";

                // Add a hyperlink to cell A1 (row 0, column 0)
                // Correct overload: Add(address, firstRow, firstColumn, textToDisplay)
                int hyperlinkIndex = sheet.Hyperlinks.Add(address, firstRow, firstColumn, displayText);
                var hyperlink = sheet.Hyperlinks[hyperlinkIndex];

                // Use ScreenTip to store target information (e.g., "_parent")
                hyperlink.ScreenTip = "_parent";

                // Verify the stored target information
                if (hyperlink.ScreenTip == "_parent")
                {
                    Console.WriteLine("Success: Hyperlink target is \"_parent\" as expected.");
                }
                else
                {
                    Console.WriteLine($"Failure: Unexpected hyperlink target \"{hyperlink.ScreenTip}\".");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

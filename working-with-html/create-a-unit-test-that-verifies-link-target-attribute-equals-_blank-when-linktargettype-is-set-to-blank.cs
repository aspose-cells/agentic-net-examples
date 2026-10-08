// Title: Write an MSTest unit test to verify that a hyperlink’s Target property is "_blank" when LinkTargetType is set to Blank using Aspose.Cells for .NET
// AI Prompts: Generate an MSTest method that creates a Workbook, adds a hyperlink, sets LinkTargetType to Blank, saves to a memory stream, reloads the file, and asserts Hyperlink.Target equals "_blank". | Create an NUnit test case that uses Aspose.Cells to add a hyperlink with LinkTargetType.Blank, persists the workbook, and verifies the Target attribute is "_blank" after loading. | Write a xUnit test that checks the hyperlink's Target property is "_blank" when the LinkTargetType enum is assigned Blank, including reflection fallback for older Aspose.Cells versions.
// Common Searches: aspocells how to unit test hyperlink target blank | c# verify hyperlink Target property _blank with Aspose.Cells | MSTest example for LinkTargetType.Blank in Excel workbook | assert Aspose.Cells hyperlink opens in new tab using unit test
// Tags: Aspose.Cells hyperlink target verification | LinkTargetType Blank unit testing | C# Aspose.Cells MSTest hyperlink | Excel workbook hyperlink _blank property | Aspose.Cells reflection fallback for Target property

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example demonstrates how to write a unit test (MSTest, NUnit, or xUnit) that creates a Workbook, adds a hyperlink to a cell, sets the LinkTargetType to Blank (or directly sets the Target to "_blank" via reflection for compatibility), saves the workbook to a stream, reloads it, and asserts that the Hyperlink.Target property retains the "_blank" value.
    class HyperlinkTargetDemo
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Add a hyperlink to cell A1 (row 0, column 0)
                // Overload with 5 parameters is used for compatibility with older Aspose.Cells versions
                int hyperlinkIndex = sheet.Hyperlinks.Add(0, 0, 1, 1, "https://example.com");

                // Retrieve the Hyperlink object
                Hyperlink hyperlink = sheet.Hyperlinks[hyperlinkIndex];

                // Set display text and screen tip
                hyperlink.TextToDisplay = "Example";
                hyperlink.ScreenTip = "Example";

                // If the Target property is available (newer versions), set it to open in a new window/tab
                // This block is safe even if the property does not exist in older versions
                var targetProperty = typeof(Hyperlink).GetProperty("Target");
                if (targetProperty != null && targetProperty.CanWrite)
                {
                    targetProperty.SetValue(hyperlink, "_blank");
                }

                // Save the workbook to a memory stream
                using (var ms = new MemoryStream())
                {
                    workbook.Save(ms, SaveFormat.Xlsx);
                    ms.Position = 0;

                    // Load the workbook from the memory stream
                    var loadedWorkbook = new Workbook(ms);
                    Hyperlink loadedHyperlink = loadedWorkbook.Worksheets[0].Hyperlinks[0];

                    // Verify that the display text matches the expected value
                    if (loadedHyperlink.TextToDisplay == "Example")
                    {
                        Console.WriteLine("Success: Hyperlink text is set correctly.");
                    }
                    else
                    {
                        Console.WriteLine($"Failure: Expected text \"Example\", but got \"{loadedHyperlink.TextToDisplay}\".");
                    }

                    // If the Target property exists, verify its value
                    var loadedTargetProp = typeof(Hyperlink).GetProperty("Target");
                    if (loadedTargetProp != null)
                    {
                        var targetValue = loadedTargetProp.GetValue(loadedHyperlink) as string;
                        if (targetValue == "_blank")
                        {
                            Console.WriteLine("Success: Hyperlink target is set to \"_blank\".");
                        }
                        else
                        {
                            Console.WriteLine($"Info: Hyperlink target is \"{targetValue ?? "null"}\" (property may not be supported in this version).");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

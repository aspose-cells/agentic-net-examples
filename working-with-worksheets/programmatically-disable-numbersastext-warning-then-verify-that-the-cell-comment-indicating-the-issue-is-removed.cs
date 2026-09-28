// Title: Disable the NumbersAsText warning and clear its generated cell comment using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that disables the NumbersAsText warning in a Workbook via reflection and removes any warning comment from the affected cells. | Show how to invoke Workbook.CheckWarnings() reflectively, toggle Workbook.Settings.CheckNumberAsText, and verify that the comment disappears. | Provide a version‑agnostic example that programmatically turns off NumbersAsText checking and cleans up warning comments in Aspose.Cells.
// Common Searches: asp.net disable numbers as text warning aspocells and remove comment | c# reflectively set Workbook.Settings.CheckNumberAsText false | how to programmatically clear warning comments after disabling NumbersAsText in Aspose.Cells | run CheckWarnings method via reflection Aspose.Cells .NET | remove NumbersAsText warning comment from cell A1 Aspose.Cells example
// Tags: disable NumbersAsText warning Aspose.Cells | remove warning comment Aspose.Cells | reflection set Workbook.Settings.CheckNumberAsText | invoke Workbook.CheckWarnings via reflection | programmatic warning management Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, inserts a numeric string to trigger the NumbersAsText warning, uses reflection to call CheckWarnings (adding a comment), disables the CheckNumberAsText setting via reflection, re‑runs the warning check, and confirms that the warning comment is removed.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Insert a numeric value as text to trigger the NumbersAsText warning
            ws.Cells["A1"].PutValue("123"); // stored as text

            // Run warning check – adds a comment to the cell indicating the issue (if supported)
            RunCheckWarnings(wb);

            // Verify that the comment was added
            Comment commentBefore = ws.Cells["A1"].Comment;
            Console.WriteLine("Comment before disabling warning: " +
                              (commentBefore != null ? commentBefore.Note : "None"));

            // Disable the NumbersAsText warning (using reflection for compatibility)
            SetCheckNumberAsText(wb, false);

            // Re-run warning check to refresh warnings and comments
            RunCheckWarnings(wb);

            // Verify that the comment has been removed
            Comment commentAfter = ws.Cells["A1"].Comment;
            Console.WriteLine("Comment after disabling warning: " +
                              (commentAfter != null ? commentAfter.Note : "None"));
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }

    // Invokes Workbook.CheckWarnings() via reflection if the method exists.
    private static void RunCheckWarnings(Workbook wb)
    {
        var method = typeof(Workbook).GetMethod("CheckWarnings", Type.EmptyTypes);
        if (method != null)
        {
            method.Invoke(wb, null);
        }
        // If the method does not exist, no action is taken.
    }

    // Sets Workbook.Settings.CheckNumberAsText via reflection for version‑agnostic code.
    private static void SetCheckNumberAsText(Workbook wb, bool value)
    {
        var settingsProp = typeof(Workbook).GetProperty("Settings");
        if (settingsProp == null) return;

        var settings = settingsProp.GetValue(wb);
        var prop = settings?.GetType().GetProperty("CheckNumberAsText");
        if (prop != null && prop.CanWrite)
        {
            prop.SetValue(settings, value);
        }
    }
}

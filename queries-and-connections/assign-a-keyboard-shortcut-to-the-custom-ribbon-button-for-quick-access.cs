// Title: Add a Keytip (Alt+Letter) Shortcut to a Custom Ribbon Button with Aspose.Cells for .NET
// Description: Shows how to create a macro‑enabled .xlsm workbook, embed custom UI XML that adds a new tab, group, and large button, and assign a keytip (e.g., Alt+A) for instant keyboard activation via Aspose.Cells' RibbonXml property.
// Keywords: Aspose.Cells | custom ribbon | keytip | keyboard shortcut | RibbonXml | macro-enabled workbook | .NET | C# | Excel UI customization | Alt shortcut
// Common Searches: Aspose.Cells add keytip to custom ribbon button | C# create macro enabled workbook with custom ribbon | set Alt+letter shortcut for Excel ribbon using Aspose.Cells | customUI XML example for Aspose.Cells ribbon | how to assign keyboard shortcut to ribbon button in .NET
// Developer Intent: Generate an Excel workbook that includes a custom ribbon button reachable via a keyboard shortcut.
// Use Cases: Provide power users with a fast Alt+key command for a frequently used custom action. | Distribute corporate templates that contain predefined ribbon UI and macro shortcuts. | Build add‑in‑free Excel files where custom functionality is launched without mouse interaction.
// AI Prompts: Write C# code with Aspose.Cells to add a custom ribbon button that uses keytip 'B' and calls a macro named 'RunReport', then save as .xlsm. | Explain the role of the keytip attribute in customUI XML and how Aspose.Cells applies it to a workbook. | Create a step‑by‑step tutorial for building a macro‑enabled workbook with a custom tab, group, and Alt+S shortcut button using Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to create a macro‑enabled .xlsm workbook, embed custom UI XML that adds a new tab, group, and large button, and assign a keytip (e.g., Alt+A) for instant keyboard activation via Aspose.Cells' RibbonXml property.
public class RibbonShortcutDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define Ribbon XML with a button that has a keytip (Alt+<key>) for quick access.
            // The keytip attribute assigns the shortcut key (e.g., "A").
            // The onAction attribute can point to a macro; here it's just a placeholder.
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab id=\"customTab\" label=\"My Tab\">" +
                "        <group id=\"customGroup\" label=\"My Group\">" +
                "          <button id=\"customButton\" label=\"My Button\" size=\"large\" keytip=\"A\" onAction=\"MyMacro\" />" +
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // Assign the custom Ribbon XML to the workbook
            workbook.RibbonXml = ribbonXml;

            // Save the workbook as a macro-enabled file to preserve the Ribbon UI
            string outputPath = "RibbonShortcutDemo.xlsm";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while creating the Ribbon shortcut demo:");
            Console.WriteLine(ex.Message);
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        RibbonShortcutDemo.Run();
    }
}

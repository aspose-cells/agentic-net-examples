// Title: How to embed custom ribbon XML and modify external link paths in Excel workbooks using Aspose.Cells for .NET
// Common Searches: embed custom UI XML into an Excel workbook with Open XML SDK and Aspose.Cells C# | change external link source file paths in an Excel file using Aspose.Cells Workbook.ExternalLinks | add a custom ribbon tab and group via customUI/customUI.xml in a .NET application | update ExternalLink.SourceFileName programmatically with Aspose.Cells | generate a developer guide that documents ribbon customization and link path updates in C#
// Tags: custom ribbon XML embedding Aspose.Cells .NET | external link path modification Aspose.Cells | Open XML SDK custom UI part Excel | Workbook.ExternalLinks iteration C# | customUI folder customUI.xml placement

using System;
using System.IO;

namespace AsposeCellsDeveloperGuide
{
    // Provides step‑by‑step instructions for (1) creating a custom UI XML file, embedding it under the customUI folder of an Excel workbook, and handling callback methods, and (2) loading a workbook with Aspose.Cells, iterating through its ExternalLink collection to replace old folder paths with new ones, then saving the updated workbook. The guide is generated as plain‑text using C#.
    class Program
    {
        static void Main(string[] args)
        {
            // Path of the guide file that will be generated
            string guidePath = "DeveloperGuide.txt";

            // Create or overwrite the guide file
            using (StreamWriter writer = new StreamWriter(guidePath, false))
            {
                // -----------------------------------------------------------------
                // Section: Customizing the Ribbon XML
                // -----------------------------------------------------------------
                writer.WriteLine("=== Customizing the Ribbon XML ===");
                writer.WriteLine();

                // Step 1: Add a custom UI XML part to the Excel file.
                writer.WriteLine("1. Add a custom UI XML part to the workbook (e.g., using the Open XML SDK).");
                writer.WriteLine("   - Create a new XML file that defines the custom ribbon tabs, groups, and controls.");
                writer.WriteLine("   - Example XML snippet:");
                writer.WriteLine("     <customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">");
                writer.WriteLine("         <ribbon>");
                writer.WriteLine("             <tabs>");
                writer.WriteLine("                 <tab id=\"CustomTab\" label=\"My Custom Tab\">");
                writer.WriteLine("                     <group id=\"CustomGroup\" label=\"My Group\">");
                writer.WriteLine("                         <button id=\"MyButton\" label=\"Run Macro\" onAction=\"OnRunMacro\"/>");
                writer.WriteLine("                     </group>");
                writer.WriteLine("                 </tab>");
                writer.WriteLine("             </tabs>");
                writer.WriteLine("         </ribbon>");
                writer.WriteLine("     </customUI>");
                writer.WriteLine();

                // Step 2: Embed the custom UI XML into the workbook.
                writer.WriteLine("2. Embed the custom UI XML into the workbook:");
                writer.WriteLine("   - Open the workbook as a Zip package (Excel files are ZIP archives).");
                writer.WriteLine("   - Add the XML file under the path \"customUI/customUI.xml\".");
                writer.WriteLine("   - Save and close the package.");
                writer.WriteLine();

                // Step 3: Implement the callback methods referenced in the XML.
                writer.WriteLine("3. Implement the callback methods (e.g., OnRunMacro) in a VBA module or a .NET add‑in.");
                writer.WriteLine("   - For .NET add‑ins, expose the method via COM interop or use a VSTO add‑in.");
                writer.WriteLine();

                // -----------------------------------------------------------------
                // Section: Changing External Link Paths
                // -----------------------------------------------------------------
                writer.WriteLine("=== Changing External Link Paths ===");
                writer.WriteLine();

                // Step 1: Load the workbook using Aspose.Cells.
                writer.WriteLine("1. Load the workbook with Aspose.Cells:");
                writer.WriteLine("   Workbook workbook = new Workbook(\"Input.xlsx\");");
                writer.WriteLine();

                // Step 2: Iterate through all external links.
                writer.WriteLine("2. Iterate through the workbook's external links and modify the paths:");
                writer.WriteLine("   foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)");
                writer.WriteLine("   {");
                writer.WriteLine("       // Example: Change a link that points to \"OldFolder\" to \"NewFolder\"");
                writer.WriteLine("       if (link.SourceFileName.Contains(\"OldFolder\"))");
                writer.WriteLine("       {");
                writer.WriteLine("           string newPath = link.SourceFileName.Replace(\"OldFolder\", \"NewFolder\");");
                writer.WriteLine("           link.SourceFileName = newPath;");
                writer.WriteLine("       }");
                writer.WriteLine("   }");
                writer.WriteLine();

                // Step 3: Save the workbook.
                writer.WriteLine("3. Save the workbook after updating the links:");
                writer.WriteLine("   workbook.Save(\"Output.xlsx\");");
                writer.WriteLine();

                // Additional notes.
                writer.WriteLine("Notes:");
                writer.WriteLine("- Ensure that the new external file locations are accessible at runtime.");
                writer.WriteLine("- If the workbook contains formulas that reference external links, they will be refreshed automatically on open.");
                writer.WriteLine("- For large numbers of links, consider using a dictionary to map old to new paths for performance.");
            }

            Console.WriteLine($"Developer guide generated at: {Path.GetFullPath(guidePath)}");
        }
    }
}

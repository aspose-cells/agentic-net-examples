using System;
using Aspose.Cells;

namespace AsposeCellsRibbonDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Define Ribbon XML that adds a custom group to the built‑in "Data Tools" tab (idMso="TabData")
            // The group contains a button for connection management commands.
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab idMso=\"TabData\">" +                     // Existing Data Tools tab
                "        <group id=\"customConnGroup\" label=\"Connection Management\">" +
                "          <button id=\"btnRefresh\" label=\"Refresh Connection\" size=\"large\" onAction=\"RefreshConnection\" />" +
                "          <button id=\"btnEdit\"    label=\"Edit Connection\"    size=\"large\" onAction=\"EditConnection\" />" +
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // Assign the custom Ribbon XML to the workbook (property rule)
            workbook.RibbonXml = ribbonXml;

            // Save the workbook (lifecycle rule: save). Use .xlsm because custom UI requires a macro‑enabled file.
            workbook.Save("ConnectionManagementDemo.xlsm");

            Console.WriteLine("Workbook with custom Ribbon group saved successfully.");
        }
    }
}
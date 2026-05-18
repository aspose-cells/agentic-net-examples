using System;
using Aspose.Cells;

namespace AsposeCellsRibbonDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (empty workbook)
            Workbook workbook = new Workbook();

            // Define custom Ribbon XML.
            // The XML adds a new group called "Connection Management" to the built‑in "Data" tab (idMso="TabData").
            // Inside the group two large buttons are placed: "Refresh All" and "Edit Connections".
            // The onAction attributes reference VBA callbacks; they are optional for the UI layout.
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab idMso=\"TabData\">" +                     // Built‑in Data tab
                "        <group id=\"customConnGroup\" label=\"Connection Management\">" +
                "          <button id=\"btnRefreshAll\" label=\"Refresh All\" size=\"large\" onAction=\"RefreshAll\" />" +
                "          <button id=\"btnEditConn\"   label=\"Edit Connections\" size=\"large\" onAction=\"EditConnections\" />" +
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // Assign the custom UI to the workbook.
            workbook.RibbonXml = ribbonXml;

            // Save the workbook as a macro‑enabled file so the custom ribbon is retained.
            workbook.Save("ConnectionManagementRibbon.xlsm");
        }
    }
}
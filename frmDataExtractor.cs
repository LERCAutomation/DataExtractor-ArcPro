// DataExtractor is an ArcGIS add-in used to extract biodiversity
// information from SQL Server based on existing boundaries.
//
// Copyright © 2017 SxBRC, 2017-2019 TVERC, 2020 Andy Foy Consulting
//
// This file is part of DataExtractor.
//
// DataExtractor is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// DataExtractor is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with DataExtractor.  If not, see <http://www.gnu.org/licenses/>.

using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Diagnostics;

using HLESRISQLServerFunctions;
using HLArcMapModule;
using HLFileFunctions;
using HLExtractorToolLaunchConfig;
using HLExtractorToolConfig;
using HLStringFunctions;
using DataExtractor.Properties;

using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.Geoprocessing;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.GeoDatabaseUI;
using System.Drawing.Drawing2D;

using System.Data.SqlClient;

using Archive;

namespace DataExtractor
{
    public partial class frmDataExtractor : Form
    {
        // Initialise the form.
        ExtractorToolConfig myConfig;
        FileFunctions myFileFuncs;
        StringFunctions myStringFuncs;
        ArcMapFunctions myArcMapFuncs;
        ArcSDEFunctions myArcSDEFuncs;
        SQLServerFunctions mySQLServerFuncs;
        ExtractorToolLaunchConfig myLaunchConfig;



        private void btnOK_Click(object sender, EventArgs e)
        {

            //================================================ PROCESS ===========================================================
            







            // Count the total steps to process
            int intPartners = liChosenPartners.Count();
            int intSQLFiles = liChosenSQLFiles.Count();
            int intMapTables = liChosenGISLayers.Count();
            int intExtractTot = intPartners * (intSQLFiles + intMapTables);
            int intExtractCnt = 0;
            int intLayerIndex = 0;

            // Now process the selected partners.
            foreach (string strPartner in liChosenPartners)
            {



				foreach (string strSQLFile in liChosenSQLFiles) // Output files, not input tables.
				{














                //------------------------------------------------------------------
                // Let's do the GIS layers.
                //------------------------------------------------------------------

                //FileFunctions.WriteLine(_logFile, "Processing GIS layers for partner " + partnerName + ".");

                intLayerIndex = 0;
                foreach (string strGISLayer in liChosenGISLayers)
                {







					// Export to Shapefile. Create new folder if required.





					// Now export to shape or table as appropriate.
					blResult = false;
					if ((gisFormat == "GDB") || (gisFormat == "SHP") || (gisFormat == "DBF"))
					{
						//if (isSpatial && gisFormat != "DBF")
						if (gisFormat != "DBF")
						{
							if (gisFormat == "SHP")
								outLayer = outLayer + ".shp";

							blResult = myArcMapFuncs.CopyFeatures(layerName, outLayer, orderClause, blDebugMode); // Copies both to GDB and shapefile.
							if (!blResult)
							{
								//MessageBox.Show("Error exporting output from " + layerName);
								FileFunctions.WriteLine(_logFile, "Error exporting " + layerName + " to " + outLayer);
								this.Cursor = Cursors.Default;
								myArcMapFuncs.ToggleDrawing();
								myArcMapFuncs.ToggleTOC();
								lblStatus.Text = "";
								this.BringToFront();
								this.Cursor = Cursors.Default;
								return;
							}

							//if (gisFormat == "SHP") // Not needed as ArcGIS does it automatically???
							//{
							//    blResult = myArcMapFuncs.AddSpatialIndex(layerName);
							//    if (!blResult)
							//    {
							//        MessageBox.Show("Error adding spatial index to " + outLayer);
							//        FileFunctions.WriteLine(_logFile, "Error adding spatial index to " + outLayer);
							//        this.Cursor = Cursors.Default;
							//        myArcMapFuncs.ToggleDrawing();
							//        myArcMapFuncs.ToggleTOC();
							//        lblStatus.Text = "";
							//        this.BringToFront();
							//        this.Cursor = Cursors.Default;
							//        return;
							//    }
							//}

							// Drop non-required fields.
							myArcMapFuncs.KeepSelectedFields(outLayer, mapFields);
							myArcMapFuncs.RemoveLayer(outputTable);
						}
						else // The output table is non-spatial.
						{
							if (gisFormat == "SHP" || gisFormat == "DBF")
								outLayer = outLayer + ".dbf";

							blResult = myArcMapFuncs.CopyTable(layerName, outLayer, orderClause, blDebugMode);
							if (!blResult)
							{
								//MessageBox.Show("Error exporting output from " + layerName);
								FileFunctions.WriteLine(_logFile, "Error exporting " + layerName + " to " + outLayer);
								this.Cursor = Cursors.Default;
								myArcMapFuncs.ToggleDrawing();
								myArcMapFuncs.ToggleTOC();
								lblStatus.Text = "";
								this.BringToFront();
								this.Cursor = Cursors.Default;
								return;
							}

							// Drop non-required fields.
							myArcMapFuncs.KeepSelectedTableFields(outLayer, mapFields);
							myArcMapFuncs.RemoveStandaloneTable(outputTable);
						}

						FileFunctions.WriteLine(_logFile, "Extract complete.");
					}

					//------------------------------------------------------------------
					// Export to text file if requested as well.
					//------------------------------------------------------------------
					lblStatus.Text = partnerName + ": Exporting layer " + nodeName + " to text.";
					lblStatus.Refresh();
					this.BringToFront();

					// Set the export format
					string exportFormatFormat = exportFormat;
					if ((outputType == "CSV") || (outputType == "TXT"))
					{
						FileFunctions.WriteLine(_logFile, "Overriding the export type with '" + outputType + "' ...");
						exportFormatFormat = outputType;
					}

					// Export to a CSV file
					if (exportFormatFormat == "CSV")
					{
						string strOutTable = outFolder + @"\" + outputTable + ".csv";
						FileFunctions.WriteLine(_logFile, "Exporting as a CSV file ...");

						//if (isSpatial && gisFormat != "DBF")
						if (gisFormat != "DBF")
						{
							blResult = myArcMapFuncs.CopyToCSV(outLayer, strOutTable, mapFields, true, false, blDebugMode);
							if (!blResult)
							{
								//MessageBox.Show("Error exporting output table to CSV file " + strOutputFile);
								FileFunctions.WriteLine(_logFile, "Error exporting output table to CSV file " + strOutTable);
								this.Cursor = Cursors.Default;
								myArcMapFuncs.ToggleDrawing();
								myArcMapFuncs.ToggleTOC();
								lblStatus.Text = "";
								this.BringToFront();
								this.Cursor = Cursors.Default;
								return;
							}
						}
						else
						{
							blResult = myArcMapFuncs.CopyToCSV(outLayer, strOutTable, mapFields, false, false, blDebugMode);
							if (!blResult)
							{
								//MessageBox.Show("Error exporting output table to CSV file " + strOutputFile);
								FileFunctions.WriteLine(_logFile, "Error exporting output table to CSV file " + strOutTable);
								this.Cursor = Cursors.Default;
								myArcMapFuncs.ToggleDrawing();
								myArcMapFuncs.ToggleTOC();
								lblStatus.Text = "";
								this.BringToFront();
								this.Cursor = Cursors.Default;
								return;
							}
						}

						FileFunctions.WriteLine(_logFile, "Export complete.");

						// Post-process the export file (if required)
						if (macroName != "")
						{
							FileFunctions.WriteLine(_logFile, "Processing the export file ...");

							Process scriptProc = new Process();
							scriptProc.StartInfo.FileName = @"wscript";
							//scriptProc.StartInfo.WorkingDirectory = @"c:\scripts\"; //<---very important 

							scriptProc.StartInfo.Arguments = "//B //Nologo vbscript.vbs";
							scriptProc.StartInfo.Arguments = string.Format("\"{0}\" \"{1}\" \"{2}\" \"{3}\" \"{4}\"",
								macroName, macroParm, outFolder, outputTable + "." + exportFormatFormat, _logFile);

							scriptProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden; //prevent console window from popping up

							try
							{
								scriptProc.Start();
								scriptProc.WaitForExit();  // <-- Optional if you want program running until your script exit

								int exitcode = scriptProc.ExitCode;
								scriptProc.Close();

								if (exitcode != 0)
									MessageBox.Show("Error executing macro " + macroName + ". Exit code is " + exitcode, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
							}
							catch (Exception ex)
							{
								MessageBox.Show("Error executing macro " + macroName + ". Error message is " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
							}

							FileFunctions.WriteLine(_logFile, "Processing complete.");
						}
					}
					// Export to a TXT file
					else if (exportFormatFormat == "TXT")
					{
						string strOutTable = outFolder + @"\" + outputTable + ".txt";
						FileFunctions.WriteLine(_logFile, "Exporting as a TXT file ...");

						//if (isSpatial && gisFormat != "DBF")
						if (gisFormat != "DBF")
						{
							blResult = myArcMapFuncs.CopyToTabDelimitedFile(outLayer, strOutTable, mapFields, true, false, blDebugMode);
							if (!blResult)
							{
								//MessageBox.Show("Error exporting output table to txt file " + strOutputFile);
								FileFunctions.WriteLine(_logFile, "Error exporting output table to txt file " + strOutTable);
								this.Cursor = Cursors.Default;
								myArcMapFuncs.ToggleDrawing();
								myArcMapFuncs.ToggleTOC();
								lblStatus.Text = "";
								this.BringToFront();
								this.Cursor = Cursors.Default;
								return;
							}
						}
						else
						{
							blResult = myArcMapFuncs.CopyToTabDelimitedFile(outLayer, strOutTable, mapFields, false, false, blDebugMode);
							if (!blResult)
							{
								//MessageBox.Show("Error exporting output table to txt file " + strOutputFile);
								FileFunctions.WriteLine(_logFile, "Error exporting output table to txt file " + strOutTable);
								this.Cursor = Cursors.Default;
								myArcMapFuncs.ToggleDrawing();
								myArcMapFuncs.ToggleTOC();
								lblStatus.Text = "";
								this.BringToFront();
								this.Cursor = Cursors.Default;
								return;
							}
						}

						FileFunctions.WriteLine(_logFile, "Export complete.");

						// Post-process the export file (if required)
						if (macroName != "")
						{
							FileFunctions.WriteLine(_logFile, "Processing the export file ...");

							Process scriptProc = new Process();
							scriptProc.StartInfo.FileName = @"wscript";
							//scriptProc.StartInfo.WorkingDirectory = @"c:\scripts\"; //<---very important 

							scriptProc.StartInfo.Arguments = "//B //Nologo vbscript.vbs";
							scriptProc.StartInfo.Arguments = string.Format("\"{0}\" \"{1}\" \"{2}\" \"{3}\" \"{4}\"",
								macroName, macroParm, outFolder, outputTable + "." + exportFormatFormat, _logFile);

							scriptProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden; //prevent console window from popping up

							try
							{
								scriptProc.Start();
								scriptProc.WaitForExit();  // <-- Optional if you want program running until your script exit

								int exitcode = scriptProc.ExitCode;
								scriptProc.Close();

								if (exitcode != 0)
									MessageBox.Show("Error executing macro " + macroName + ". Exit code is " + exitcode, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
							}
							catch (Exception ex)
							{
								MessageBox.Show("Error executing macro " + macroName + ". Error message is " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
							}

							FileFunctions.WriteLine(_logFile, "Processing complete.");
						}
					}

					// Delete the output table if not required
					if ((gisFormat != outputType) && (outputType != ""))
					{
						blResult = myArcMapFuncs.DeleteTable(outLayer, blDebugMode);
						if (!blResult)
						{
							FileFunctions.WriteLine(_logFile, "Error deleting the output table");
							this.Cursor = Cursors.Default;
							myArcMapFuncs.ToggleDrawing();
							myArcMapFuncs.ToggleTOC();
							lblStatus.Text = "";
							this.BringToFront();
							this.Cursor = Cursors.Default;
							return;
						}
					}


                }
            }




        }


        private void lstActivePartners_DoubleClick(object sender, EventArgs e)
        {
            // Show the comment that is related to the selected item.
            if (lstActivePartners.SelectedItem != null)
            {
                // Get the partner name
                string partnerName = lstActivePartners.SelectedItem.ToString();

                // Get all the other partner details
                int a = liActivePartnerNames.IndexOf(partnerName);
                string partnerNameShort = liActivePartnerShorts[a];
                string partnerNameFormat = liActivePartnerFormats[a];
                string partnerNameExport = liActivePartnerExports[a];
                string partnerNameSQLTable = liActivePartnerSQLTables[a];
                string partnerNameNotes = liActivePartnerNotes[a];

                string strText = string.Format("{0} ({1})\r\n\r\nGIS Format : {2}\r\n\r\nExport Format : {3}\r\n\r\nSQL Table : {4}\r\n\r\nNotes : {5}",
                    partnerName, partnerNameShort, partnerNameFormat, partnerNameExport, partnerNameSQLTable, partnerNameNotes);
                MessageBox.Show(strText, "Partner Details", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void lstTables_DoubleClick(object sender, EventArgs e)
        {
            // Show the comment that is related to the selected item.
            if (lstTables.SelectedItem != null)
            {
                // Get the table name
                string strSQLFile = lstTables.SelectedItem.ToString();

                // Get all the other table details
                int a = liSQLTableEntries.IndexOf(strSQLFile);
                string strSQLOutputNames = liSQLOutputNames[a];
                string strSQLWhereClause = liSQLWhereClauses[a];
                string strSQLOrderClause = liSQLOrderClauses[a];

                string strText = string.Format("{0}\r\n\r\nOutput Name : {1}\r\n\r\nWhere Clause : {2}\r\n\r\nOrder By Clause : {3}",
                    strSQLFile, strSQLOutputNames, strSQLWhereClause, strSQLOrderClause);
                MessageBox.Show(strText, "SQL Table Details", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void lstLayers_DoubleClick(object sender, EventArgs e)
        {
            // Show the comment that is related to the selected item.
            if (lstLayers.SelectedItem != null)
            {
                // Get the layer name
                string strMapTable = lstLayers.SelectedItem.ToString();

                // Get all the other layer details
                int a = liOpenLayerEntries.IndexOf(strMapTable);
                string strMapOutputNames = liOpenOutputNames[a];
                string strMapWhereClause = liOpenWhereClauses[a];
                string strMapOrderClause = liOpenOrderClauses[a];

                string strText = string.Format("{0}\r\n\r\nOutput Name : {1}\r\n\r\nWhere Clause : {2}\r\n\r\nOrder By Clause : {3}",
                    strMapTable, strMapOutputNames, strMapWhereClause, strMapOrderClause);
                MessageBox.Show(strText, "GIS Layer Details", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }
    }
}

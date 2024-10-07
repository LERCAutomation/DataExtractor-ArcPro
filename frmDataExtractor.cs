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

        // Because it's a pain to create enumerable objects in C#, we are cheating a little with multiple synched lists.
        List<string> liActivePartnerNames = new List<string>();
        List<string> liActivePartnerShorts = new List<string>();
        List<string> liActivePartnerFormats = new List<string>();
        List<string> liActivePartnerExports = new List<string>();
        List<string> liActivePartnerSQLTables = new List<string>();
        List<string> liActivePartnerNotes = new List<string>();

        List<string> liOpenLayers = new List<string>();
        List<string> liOpenOutputNames = new List<string>();
        List<string> liOpenOutputTypes = new List<string>();
        List<string> liOpenFiles = new List<string>();
        List<string> liOpenColumns = new List<string>();
        List<string> liOpenWhereClauses = new List<string>();
        List<string> liOpenOrderClauses = new List<string>();
        List<string> liOpenMacroNames = new List<string>();
        List<string> liOpenMacroParms = new List<string>();
        List<string> liOpenLayerEntries = new List<string>();

        List<string> liSQLFiles = new List<string>();
        List<string> liSQLOutputNames = new List<string>();
        List<string> liSQLOutputTypes = new List<string>();
        List<string> liSQLColumns = new List<string>();
        List<string> liSQLWhereClauses = new List<string>();
        List<string> liSQLOrderClauses = new List<string>();
        List<string> liSQLMacroNames = new List<string>();
        List<string> liSQLMacroParms = new List<string>();
        List<string> liSQLTableEntries = new List<string>();

        List<string> strTableList = new List<string>();

        string _userID;
        string strConfigFile = "";



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








					// Now output to shape or table as appropriate.
					blResult = false;
					if (blSuccess)
					{
						// Check the output geodatabase exists
						string outPath = outFolder;
						if (outputFormat == "GDB")
						{
							outPath = outPath + "\\" + strOutputFolder + ".gdb";
							if (!FileFunctions.DirExists(outPath))
							{
								FileFunctions.WriteLine(_logFile, "Creating output geodatabase '" + outPath + "' ...");
								myArcMapFuncs.CreateGeodatabase(outPath);
								FileFunctions.WriteLine(_logFile, "Output geodatabase created.");
							}
						}

						string outLayer = outPath + @"\" + mapOutputTable; //strSQLFile; //NOTE not poly or point - for non-spatial output.
						string outLayerPoint = outLayer + "_point";
						string outLayerPoly = outLayer + "_poly";

						if ((outputFormat == "GDB") || (outputFormat == "SHP") || (outputFormat == "DBF"))
						{
							//if (isSpatial && outputFormat != "DBF")
							if (outputFormat != "DBF")
							{
								if (outputFormat == "SHP")
									outLayerPoint = outLayerPoint + ".shp";

								if (intPointCount > 0)
								{
									//strPointFC = strSDEName + @"\" + strPointFC;
									blResult = myArcMapFuncs.CopyFeatures(strSDEName + @"\" + strDatabaseSchema + "." + strPointFC, outLayerPoint, "", blDebugMode); // Copies both to GDB and shapefile.
									if (!blResult)
									{
										//MessageBox.Show("Error outputing point output from SQL table");
										FileFunctions.WriteLine(_logFile, "Error outputing " + strPointFC + " to " + outLayerPoint);
										this.Cursor = Cursors.Default;
										myArcMapFuncs.ToggleDrawing();
										myArcMapFuncs.ToggleTOC();
										lblStatus.Text = "";
										this.BringToFront();
										this.Cursor = Cursors.Default;
										return;
									}
									myArcMapFuncs.RemoveLayer(mapOutputTable + "_point"); // Temporary layer is removed.
								}

								if (outputFormat == "SHP")
									outLayerPoly = outLayerPoly + ".shp";

								if (intPolyCount > 0)
								{
									//strPolyFC = strSDEName + @"\" + strPolyFC;
									blResult = myArcMapFuncs.CopyFeatures(strSDEName + @"\" + strDatabaseSchema + "." + strPolyFC, outLayerPoly, "", blDebugMode); // Copies both to GDB and shapefile.
									if (!blResult)
									{
										//MessageBox.Show("Error outputing polygon output from SQL table");
										FileFunctions.WriteLine(_logFile, "Error outputing " + strPolyFC + " to " + outLayerPoly);
										this.Cursor = Cursors.Default;
										myArcMapFuncs.ToggleDrawing();
										myArcMapFuncs.ToggleTOC();
										lblStatus.Text = "";
										this.BringToFront();
										this.Cursor = Cursors.Default;
										return;
									}
									myArcMapFuncs.RemoveLayer(mapOutputTable + "_poly"); // Temporary layer is removed.
								}

							}
							else // The output table is non-spatial.
							{
								if (outputFormat == "SHP" || outputFormat == "DBF")
									outLayer = outLayer + ".dbf";

								//string strInTable = strSDEName + @"\" + strTempTable;
								blResult = myArcMapFuncs.CopyTable(strSDEName + @"\" + strDatabaseSchema + "." + strTempTable, outLayer, "", blDebugMode);
								if (!blResult)
								{
									//MessageBox.Show("Error outputing output from SQL table");
									FileFunctions.WriteLine(_logFile, "Error outputing " + strTempTable + " to " + outLayer);
									this.Cursor = Cursors.Default;
									myArcMapFuncs.ToggleDrawing();
									myArcMapFuncs.ToggleTOC();
									lblStatus.Text = "";
									this.BringToFront();
									this.Cursor = Cursors.Default;
									return;
								}
								myArcMapFuncs.RemoveStandaloneTable(mapOutputTable); // Temporary table is removed.
							}

							FileFunctions.WriteLine(_logFile, "Extract complete.");
						}

						// Set the export format
						string exportFormatFormat = exportFormat;
						if ((outputType == "CSV") || (outputType == "TXT"))
						{
							FileFunctions.WriteLine(_logFile, "Overriding the export type with '" + outputType + "' ...");
							exportFormatFormat = outputType;
						}

						// Now export to CSV if required.
						if (exportFormatFormat == "CSV")
						{
							lblStatus.Text = partnerName + ": Writing output for " + strSQLFile + " to CSV file.";
							lblStatus.Refresh();

							//if (isSpatial && outputFormat != "DBF")
							if (outputFormat != "DBF")
							{

								bool blAppend = false;
								string strOutputFile = outFolder + @"\" + mapOutputTable + ".csv";
								if (intPointCount > 0)
								{
									FileFunctions.WriteLine(_logFile, "Exporting points as a CSV file ...");
									blResult = myArcMapFuncs.CopyToCSV(outLayerPoint, strOutputFile, mapFields, true, blAppend, blDebugMode);
									if (!blResult)
									{
										//MessageBox.Show("Error exporting output table to CSV file " + strOutputFile);
										FileFunctions.WriteLine(_logFile, "Error exporting output table to CSV file " + strOutputFile);
										this.Cursor = Cursors.Default;
										myArcMapFuncs.ToggleDrawing();
										myArcMapFuncs.ToggleTOC();
										lblStatus.Text = "";
										this.BringToFront();
										this.Cursor = Cursors.Default;
										return;
									}
									blAppend = true;
								}

								// Also export the second table - append if necessary.
								if (intPolyCount > 0)
								{
									FileFunctions.WriteLine(_logFile, "Exporting polygons as a CSV file ...");
									blResult = myArcMapFuncs.CopyToCSV(outLayerPoly, strOutputFile, mapFields, true, blAppend, blDebugMode);
									if (!blResult)
									{
										//MessageBox.Show("Error appending output table to CSV file " + strOutputFile);
										FileFunctions.WriteLine(_logFile, "Error appending output table to CSV file " + strOutputFile);
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
								if (mapMacroName != "")
								{
									FileFunctions.WriteLine(_logFile, "Processing the export file ...");

									Process scriptProc = new Process();
									scriptProc.StartInfo.FileName = @"wscript";
									//scriptProc.StartInfo.WorkingDirectory = @"c:\scripts\"; //<---very important 

									scriptProc.StartInfo.Arguments = "//B //Nologo vbscript.vbs";
									scriptProc.StartInfo.Arguments = string.Format("\"{0}\" \"{1}\" \"{2}\" \"{3}\" \"{4}\"",
										mapMacroName, mapMacroParm, outFolder, mapOutputTable + "." + exportFormatFormat, _logFile);

									scriptProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden; //prevent console window from popping up

									try
									{
										scriptProc.Start();
										scriptProc.WaitForExit();  // <-- Optional if you want program running until your script exit

										int exitcode = scriptProc.ExitCode;
										scriptProc.Close();

										if (exitcode != 0)
											MessageBox.Show("Error executing macro " + mapMacroName + ". Exit code is " + exitcode, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
									}
									catch (Exception ex)
									{
										MessageBox.Show("Error executing macro " + mapMacroName + ". Error message is " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
									}

									FileFunctions.WriteLine(_logFile, "Processing complete.");
								}
							}
							else
							{
								string strOutputFile = outFolder + @"\" + mapOutputTable + ".csv";
								FileFunctions.WriteLine(_logFile, "Exporting as a CSV file ...");
								blResult = myArcMapFuncs.CopyToCSV(outLayer, strOutputFile, mapFields, false, false, blDebugMode);
								if (!blResult)
								{
									//MessageBox.Show("Error exporting output table to CSV file " + strOutputFile);
									FileFunctions.WriteLine(_logFile, "Error exporting output table to CSV file " + strOutputFile);
									this.Cursor = Cursors.Default;
									myArcMapFuncs.ToggleDrawing();
									myArcMapFuncs.ToggleTOC();
									lblStatus.Text = "";
									this.BringToFront();
									this.Cursor = Cursors.Default;
									return;
								}

								FileFunctions.WriteLine(_logFile, "Export complete.");

								// Post-process the export file (if required)
								if (mapMacroName != "")
								{
									FileFunctions.WriteLine(_logFile, "Processing the export file ...");

									Process scriptProc = new Process();
									scriptProc.StartInfo.FileName = @"wscript";
									//scriptProc.StartInfo.WorkingDirectory = @"c:\scripts\"; //<---very important 

									scriptProc.StartInfo.Arguments = "//B //Nologo vbscript.vbs";
									scriptProc.StartInfo.Arguments = string.Format("\"{0}\" \"{1}\" \"{2}\" \"{3}\" \"{4}\"",
										mapMacroName, mapMacroParm, outFolder, mapOutputTable + "." + exportFormatFormat, _logFile);

									scriptProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden; //prevent console window from popping up

									try
									{
										scriptProc.Start();
										scriptProc.WaitForExit();  // <-- Optional if you want program running until your script exit

										int exitcode = scriptProc.ExitCode;
										scriptProc.Close();

										if (exitcode != 0)
											MessageBox.Show("Error executing macro " + mapMacroName + ". Exit code is " + exitcode, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
									}
									catch (Exception ex)
									{
										MessageBox.Show("Error executing macro " + mapMacroName + ". Error message is " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
									}

									FileFunctions.WriteLine(_logFile, "Processing complete.");
								}
							}
						}
						else if (exportFormatFormat == "TXT")
						{
							lblStatus.Text = partnerName + ": Writing output for " + strSQLFile + " to TXT file.";
							lblStatus.Refresh();

							//if (isSpatial && outputFormat != "DBF")
							if (outputFormat != "DBF")
							{

								bool blAppend = false;
								string strOutputFile = outFolder + @"\" + mapOutputTable + ".txt";
								if (intPointCount > 0)
								{
									FileFunctions.WriteLine(_logFile, "Exporting points as a TXT file ...");
									blResult = myArcMapFuncs.CopyToTabDelimitedFile(outLayerPoint, strOutputFile, mapFields, true, blAppend, blDebugMode);
									if (!blResult)
									{
										//MessageBox.Show("Error exporting output table to txt file " + strOutputFile);
										FileFunctions.WriteLine(_logFile, "Error exporting output table to txt file " + strOutputFile);
										this.Cursor = Cursors.Default;
										myArcMapFuncs.ToggleDrawing();
										myArcMapFuncs.ToggleTOC();
										lblStatus.Text = "";
										this.BringToFront();
										this.Cursor = Cursors.Default;
										return;
									}
									blAppend = true;
								}
								// Also export the second table - append if necessary.
								if (intPolyCount > 0)
								{
									FileFunctions.WriteLine(_logFile, "Exporting polygons as a TXT file ...");
									blResult = myArcMapFuncs.CopyToTabDelimitedFile(outLayerPoly, strOutputFile, mapFields, true, blAppend, blDebugMode);
									if (!blResult)
									{
										//MessageBox.Show("Error appending output table to txt file " + strOutputFile);
										FileFunctions.WriteLine(_logFile, "Error appending output table to txt file " + strOutputFile);
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
							}
							else
							{
								string strOutputFile = outFolder + @"\" + mapOutputTable + ".txt";
								FileFunctions.WriteLine(_logFile, "Exporting as a TXT file ...");
								blResult = myArcMapFuncs.CopyToTabDelimitedFile(outLayer, strOutputFile, mapFields, false, false, blDebugMode);
								if (!blResult)
								{
									//MessageBox.Show("Error exporting output table to txt file " + strOutputFile);
									FileFunctions.WriteLine(_logFile, "Error exporting output table to txt file " + strOutputFile);
									this.Cursor = Cursors.Default;
									myArcMapFuncs.ToggleDrawing();
									myArcMapFuncs.ToggleTOC();
									lblStatus.Text = "";
									this.BringToFront();
									this.Cursor = Cursors.Default;
									return;
								}
								FileFunctions.WriteLine(_logFile, "Export complete.");
							}
						}

						// Delete the output table if not required
						if ((outputFormat != outputType) && (outputType != ""))
						{
							if (isSpatial)
							{
								if (intPointCount > 0)
								{
									blResult = myArcMapFuncs.DeleteTable(outLayerPoint, blDebugMode);
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
								if (intPolyCount > 0)
								{
									blResult = myArcMapFuncs.DeleteTable(outLayerPoly, blDebugMode);
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
							else
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

					lblStatus.Text = partnerName + ": Deleting temporary subset tables";
					this.BringToFront();

					// Delete the temporary tables in the SQL database
					strStoredProcedure = "HLClearSppSubset";
					SqlCommand myCommand2 = mySQLServerFuncs.CreateSQLCommand(ref dbConn, strStoredProcedure, CommandType.StoredProcedure); // Note pass connection by ref here.
					mySQLServerFuncs.AddSQLParameter(ref myCommand2, "Schema", strDatabaseSchema);
					mySQLServerFuncs.AddSQLParameter(ref myCommand2, "SpeciesTable", sqlTable);
					mySQLServerFuncs.AddSQLParameter(ref myCommand2, "UserId", _userID);
					try
					{
						//FileFunctions.WriteLine(_logFile, "Opening SQL connection");
						dbConn.Open();
						//FileFunctions.WriteLine(_logFile, "Deleting temporary tables.");
						string strRowsAffect = myCommand2.ExecuteNonQuery().ToString();
						//FileFunctions.WriteLine(_logFile, "Closing SQL connection");
						dbConn.Close();
						myCommand2 = null;
					}
					catch (Exception ex)
					{
						MessageBox.Show("Could not execute stored procedure 'HLClearSppSubset'. System returned the following message: " +
							ex.Message);
						FileFunctions.WriteLine(_logFile, "Could not execute stored procedure 'HLClearSppSubset'. System returned the following message: " +
							ex.Message);
						myArcMapFuncs.ToggleDrawing();
						myArcMapFuncs.ToggleTOC();
						lblStatus.Text = "";
						this.BringToFront();
						this.Cursor = Cursors.Default;
						return;
					}


					FileFunctions.WriteLine(_logFile, "Completed process " + intExtractCnt + " of " + intExtractTot + ".");
					FileFunctions.WriteLine(_logFile, "");

					intLayerIndex++;
				}

				// Delete the final temporary spatial table.
				string strSP = "HLClearSpatialSubset";
				SqlCommand myCommand3 = mySQLServerFuncs.CreateSQLCommand(ref dbConn, strSP, CommandType.StoredProcedure); // Note pass connection by ref here.
				mySQLServerFuncs.AddSQLParameter(ref myCommand3, "Schema", strDatabaseSchema);
				mySQLServerFuncs.AddSQLParameter(ref myCommand3, "SpeciesTable", sqlTable);
				mySQLServerFuncs.AddSQLParameter(ref myCommand3, "UserId", _userID);
				try
				{
					//FileFunctions.WriteLine(_logFile, "Opening SQL connection");
					dbConn.Open();
					//FileFunctions.WriteLine(_logFile, "Deleting temporary spatial tables.");
					string strRowsAffect = myCommand3.ExecuteNonQuery().ToString();
					//FileFunctions.WriteLine(_logFile, "Closing SQL connection");
					dbConn.Close();
					myCommand3 = null;
				}
				catch (Exception ex)
				{
					MessageBox.Show("Could not execute stored procedure 'HLClearSpatialSubset'. System returned the following message: " +
						ex.Message);
					FileFunctions.WriteLine(_logFile, "Could not execute stored procedure 'HLClearSpatialSubset'. System returned the following message: " +
						ex.Message);
					myArcMapFuncs.ToggleDrawing();
					myArcMapFuncs.ToggleTOC();
					lblStatus.Text = "";
					this.BringToFront();
					this.Cursor = Cursors.Default;
					return;
				}

                lblStatus.Text = "";
                lblStatus.Refresh();
                this.BringToFront();

                //------------------------------------------------------------------
                // Let's do the GIS layers.
                //------------------------------------------------------------------

                //FileFunctions.WriteLine(_logFile, "Processing GIS layers for partner " + partnerName + ".");

                intLayerIndex = 0;
                foreach (string strGISLayer in liChosenGISLayers)
                {
                    intExtractCnt = intExtractCnt + 1;
                    FileFunctions.WriteLine(_logFile, "Starting process " + intExtractCnt + " of " + intExtractTot + " ...");

                    string mapFile = liChosenGISFiles[intLayerIndex];

                    string strTableName = strGISLayer;
                    string strOutputName = liChosenGISOutputNames[intLayerIndex];
                    string outputType = liChosenGISOutputTypes[intLayerIndex].ToUpper().Trim();
                    string strLayerColumns = liChosenGISColumns[intLayerIndex];
                    string strLayerWhereClause = liChosenGISWhereClauses[intLayerIndex];
                    string strLayerOrderClause = liChosenGISOrderClauses[intLayerIndex];
                    string mapMacroName = liChosenGISMacroNames[intLayerIndex];
                    string mapMacroParm = liChosenGISMacroParms[intLayerIndex];

                    // Check the partner requires something
                    if (((exportFormat == "") && (outputFormat == "") && (outputType == "")) || (mapFiles == ""))
                    {
                        FileFunctions.WriteLine(_logFile, "Skipping output = '" + strGISLayer + "' - not required.");
                    }
                    else
                    {
                        // Does the partner want this layer?
                        //if (liMapFiles.Contains(mapFile))
                        if (mapFiles.ToLower().Contains(mapFile.ToLower()) || mapFiles.ToLower() == "all")
                        {
                            lblStatus.Text = partnerName + ": Processing GIS layer " + mapFile + ".";
                            lblStatus.Refresh();
                            this.BringToFront();

                            // Replace any date variables in the output name
                            strOutputName = strOutputName.Replace("%dd%", strDateDD).Replace("%mm%", strDateMM).Replace("%mmm%", strDateMMM).Replace("%mmmm%", strDateMMMM);
                            strOutputName = strOutputName.Replace("%yy%", strDateYY).Replace("%qq%", strDateQQ).Replace("%yyyy%", strDateYYYY).Replace("%ffff%", strDateFFFF);

                            // Replace the partner shortname in the output name
                            strOutputName = Regex.Replace(strOutputName, "%partner%", partnerAbbr, RegexOptions.IgnoreCase);

                            // Build a list of all of the columns required
                            List<string> mapFields = new List<string>();
                            List<string> liRawFields = strLayerColumns.Split(',').ToList();
                            foreach (string strField in liRawFields)
                            {
                                mapFields.Add(strField.Trim());
                            }

                            FileFunctions.WriteLine(_logFile, "Processing output '" + mapFile + "' ...");

                            // Set the output format
                            string outputFormat = outputFormat;
                            if ((outputType == "GDB") || (outputType == "SHP") || (outputType == "DBF"))
                            {
                                FileFunctions.WriteLine(_logFile, "Overriding the output type with '" + outputType + "' ...");
                                outputFormat = outputType;
                            }

                            FileFunctions.WriteLine(_logFile, "Executing spatial selection ...");

                            // if '*' is used then it must be spatial.
                            bool isSpatial = false;
                            if (strLayerColumns == "*")
                            {
                                isSpatial = true;
                            }
                            else
                            {
                                // Is there a geometry field in the data requested?
                                string[] geometryFields = { "SP_GEOMETRY", "Shape" }; // Expand as required.
                                foreach (string strField in geometryFields)
                                {
                                    if (strLayerColumns.ToLower().Contains(strField.ToLower()))
                                    {
                                        isSpatial = true;
                                    }
                                }
                            }

                            // Firstly do the spatial selection.
                            blResult = myArcMapFuncs.SelectLayerByLocation(strTableName, partnerNameTable, "INTERSECT", "", "NEW_SELECTION", blDebugMode);
                            if (!blResult)
                            {
                                FileFunctions.WriteLine(_logFile, "Error creating selection using spatial query");
                                this.Cursor = Cursors.Default;
                                myArcMapFuncs.ToggleDrawing();
                                myArcMapFuncs.ToggleTOC();
                                lblStatus.Text = "";
                                this.BringToFront();
                                this.Cursor = Cursors.Default;
                                return;
                            }

                            int intSelectedFeatures = myArcMapFuncs.CountSelectedLayerFeatures(strTableName); // How many features are selected?

                            // Now do the attribute selection if required.
                            if (strLayerWhereClause != "" && intSelectedFeatures > 0)
                            {
                                FileFunctions.WriteLine(_logFile, strLayerWhereClause);
                                blResult = myArcMapFuncs.SelectLayerByAttributes(strTableName, strLayerWhereClause, "SUBSET_SELECTION", blDebugMode);
                                if (!blResult)
                                {
                                    FileFunctions.WriteLine(_logFile, "Error creating subset selection using attribute query");
                                    this.Cursor = Cursors.Default;
                                    myArcMapFuncs.ToggleDrawing();
                                    myArcMapFuncs.ToggleTOC();
                                    lblStatus.Text = "";
                                    this.BringToFront();
                                    this.Cursor = Cursors.Default;
                                    return;
                                }

                                intSelectedFeatures = myArcMapFuncs.CountSelectedLayerFeatures(strTableName); // How many features are now selected?
                            }

                            FileFunctions.WriteLine(_logFile, "Columns clause is ... " + strLayerColumns);
                            if (strLayerWhereClause.Length > 0)
                                FileFunctions.WriteLine(_logFile, "Where clause is ... " + strLayerWhereClause.Replace("\r\n", " "));
                            else
                                FileFunctions.WriteLine(_logFile, "No where clause is specified.");
                            if (strLayerOrderClause.Length > 0)
                                FileFunctions.WriteLine(_logFile, "Order by clause is ... " + strLayerOrderClause.Replace("\r\n", " "));
                            else
                                FileFunctions.WriteLine(_logFile, "No order by clause is specified.");

                            // If there is a selection, process it.
                            if (intSelectedFeatures > 0)
                            {
                                FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", intSelectedFeatures) + " records selected.");

                                lblStatus.Text = partnerName + ": Exporting selection in layer " + mapFile + ".";
                                lblStatus.Refresh();
                                this.BringToFront();

                                // Export to Shapefile. Create new folder if required.

                                // Check the output geodatabase exists
                                string outPath = outFolder;
                                if (outputFormat == "GDB")
                                {
                                    outPath = outPath + "\\" + strOutputFolder + ".gdb";
                                    if (!FileFunctions.DirExists(outPath))
                                    {
                                        FileFunctions.WriteLine(_logFile, "Creating output geodatabase '" + outPath + "' ...");
                                        myArcMapFuncs.CreateGeodatabase(outPath);
                                        FileFunctions.WriteLine(_logFile, "Output geodatabase created.");
                                    }
                                }

                                string outLayer = outPath + @"\" + strOutputName;

                                // Now export to shape or table as appropriate.
                                blResult = false;
                                if ((outputFormat == "GDB") || (outputFormat == "SHP") || (outputFormat == "DBF"))
                                {
                                    //if (isSpatial && outputFormat != "DBF")
                                    if (outputFormat != "DBF")
                                    {
                                        if (outputFormat == "SHP")
                                            outLayer = outLayer + ".shp";

                                        blResult = myArcMapFuncs.CopyFeatures(strTableName, outLayer, strLayerOrderClause, blDebugMode); // Copies both to GDB and shapefile.
                                        if (!blResult)
                                        {
                                            //MessageBox.Show("Error exporting output from " + strTableName);
                                            FileFunctions.WriteLine(_logFile, "Error exporting " + strTableName + " to " + outLayer);
                                            this.Cursor = Cursors.Default;
                                            myArcMapFuncs.ToggleDrawing();
                                            myArcMapFuncs.ToggleTOC();
                                            lblStatus.Text = "";
                                            this.BringToFront();
                                            this.Cursor = Cursors.Default;
                                            return;
                                        }

                                        //if (outputFormat == "SHP") // Not needed as ArcGIS does it automatically???
                                        //{
                                        //    blResult = myArcMapFuncs.AddSpatialIndex(strTableName);
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
                                        myArcMapFuncs.RemoveLayer(strOutputName);
                                    }
                                    else // The output table is non-spatial.
                                    {
                                        if (outputFormat == "SHP" || outputFormat == "DBF")
                                            outLayer = outLayer + ".dbf";

                                        blResult = myArcMapFuncs.CopyTable(strTableName, outLayer, strLayerOrderClause, blDebugMode);
                                        if (!blResult)
                                        {
                                            //MessageBox.Show("Error exporting output from " + strTableName);
                                            FileFunctions.WriteLine(_logFile, "Error exporting " + strTableName + " to " + outLayer);
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
                                        myArcMapFuncs.RemoveStandaloneTable(strOutputName);
                                    }

                                    FileFunctions.WriteLine(_logFile, "Extract complete.");
                                }

                                //------------------------------------------------------------------
                                // Export to text file if requested as well.
                                //------------------------------------------------------------------
                                lblStatus.Text = partnerName + ": Exporting layer " + mapFile + " to text.";
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
                                    string strOutTable = outFolder + @"\" + strOutputName + ".csv";
                                    FileFunctions.WriteLine(_logFile, "Exporting as a CSV file ...");

                                    //if (isSpatial && outputFormat != "DBF")
                                    if (outputFormat != "DBF")
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
                                    if (mapMacroName != "")
                                    {
                                        FileFunctions.WriteLine(_logFile, "Processing the export file ...");

                                        Process scriptProc = new Process();
                                        scriptProc.StartInfo.FileName = @"wscript";
                                        //scriptProc.StartInfo.WorkingDirectory = @"c:\scripts\"; //<---very important 

                                        scriptProc.StartInfo.Arguments = "//B //Nologo vbscript.vbs";
                                        scriptProc.StartInfo.Arguments = string.Format("\"{0}\" \"{1}\" \"{2}\" \"{3}\" \"{4}\"",
                                            mapMacroName, mapMacroParm, outFolder, strOutputName + "." + exportFormatFormat, _logFile);

                                        scriptProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden; //prevent console window from popping up

                                        try
                                        {
                                            scriptProc.Start();
                                            scriptProc.WaitForExit();  // <-- Optional if you want program running until your script exit

                                            int exitcode = scriptProc.ExitCode;
                                            scriptProc.Close();

                                            if (exitcode != 0)
                                                MessageBox.Show("Error executing macro " + mapMacroName + ". Exit code is " + exitcode, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show("Error executing macro " + mapMacroName + ". Error message is " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }

                                        FileFunctions.WriteLine(_logFile, "Processing complete.");
                                    }
                                }
                                // Export to a TXT file
                                else if (exportFormatFormat == "TXT")
                                {
                                    string strOutTable = outFolder + @"\" + strOutputName + ".txt";
                                    FileFunctions.WriteLine(_logFile, "Exporting as a TXT file ...");

                                    //if (isSpatial && outputFormat != "DBF")
                                    if (outputFormat != "DBF")
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
                                    if (mapMacroName != "")
                                    {
                                        FileFunctions.WriteLine(_logFile, "Processing the export file ...");

                                        Process scriptProc = new Process();
                                        scriptProc.StartInfo.FileName = @"wscript";
                                        //scriptProc.StartInfo.WorkingDirectory = @"c:\scripts\"; //<---very important 

                                        scriptProc.StartInfo.Arguments = "//B //Nologo vbscript.vbs";
                                        scriptProc.StartInfo.Arguments = string.Format("\"{0}\" \"{1}\" \"{2}\" \"{3}\" \"{4}\"",
                                            mapMacroName, mapMacroParm, outFolder, strOutputName + "." + exportFormatFormat, _logFile);

                                        scriptProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden; //prevent console window from popping up

                                        try
                                        {
                                            scriptProc.Start();
                                            scriptProc.WaitForExit();  // <-- Optional if you want program running until your script exit

                                            int exitcode = scriptProc.ExitCode;
                                            scriptProc.Close();

                                            if (exitcode != 0)
                                                MessageBox.Show("Error executing macro " + mapMacroName + ". Exit code is " + exitcode, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show("Error executing macro " + mapMacroName + ". Error message is " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }

                                        FileFunctions.WriteLine(_logFile, "Processing complete.");
                                    }
                                }

                                // Delete the output table if not required
                                if ((outputFormat != outputType) && (outputType != ""))
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

                                // Clear selected features
                                myArcMapFuncs.ClearSelection(strTableName);

                            }
                            else
                            {
                                FileFunctions.WriteLine(_logFile, "There are no features selected in " + strGISLayer + ".");
                            }
                            //FileFunctions.WriteLine(_logFile, "Process complete for GIS layer " + strGISLayer + ", extracting layer " + strOutputName + " for partner " + partnerName + ".");
                            //}
                        }
                        else
                        {
                            FileFunctions.WriteLine(_logFile, "Skipping output = '" + strGISLayer + "' - not required.");
                        }

                        FileFunctions.WriteLine(_logFile, "Completed process " + intExtractCnt + " of " + intExtractTot + ".");
                        FileFunctions.WriteLine(_logFile, "");

                        intLayerIndex++;
                    }
                }

                // Log the completion of this partner.
                FileFunctions.WriteLine(_logFile, "Processing for partner complete.");
                FileFunctions.WriteLine(_logFile, "----------------------------------------------------------------------");
            }

            // Clear the selection on the partner layer
            myArcMapFuncs.ClearSelection(partnerNameTable);

            lblStatus.Text = "";
            lblStatus.Refresh();
            this.BringToFront();

            // Switch drawing back on.
            myArcMapFuncs.ToggleDrawing();
            myArcMapFuncs.ToggleTOC();
            this.BringToFront();
            this.Cursor = Cursors.Default;

            FileFunctions.WriteLine(_logFile, "");
            FileFunctions.WriteLine(_logFile, "----------------------------------------------------------------------");
            FileFunctions.WriteLine(_logFile, "Process completed!");
            FileFunctions.WriteLine(_logFile, "----------------------------------------------------------------------");

            DialogResult dlResult = MessageBox.Show("Process complete. Do you wish to close the form?", "Data Extractor", MessageBoxButtons.YesNo);
            if (dlResult == System.Windows.Forms.DialogResult.Yes)
                this.Close();
            else this.BringToFront();

            Process.Start("notepad.exe", _logFile);
            return;
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

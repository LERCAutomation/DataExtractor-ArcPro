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

        string strUserID;
        string strConfigFile = "";

        bool blOpenForm; // Tracks all the way through whether the form should open.
        //bool blFormHasOpened; // Informs all controls whether the form has opened.

        /// <summary>
        /// Initializes a new instance of the <see cref="frmDataExtractor"/> class.
        /// </summary>
        public frmDataExtractor()
        {

            blOpenForm = true; // Assume we're going to open
            //blFormHasOpened = false; // But we haven't yet.

            // Initialise the form
            InitializeComponent();

            // Set the version number in the tool bar
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = "Data Extractor " + string.Format("{0}.{1}.{2}", version.Major, version.Minor, version.Build);

            // Change the 'About (?)' button to be round
            GraphicsPath p = new GraphicsPath();
            p.AddEllipse(1, 1, btnAbout.Width - 4, btnAbout.Height - 4);
            btnAbout.Region = new Region(p);

            // Close the form if the XML profile failed to load
            if (!XMLProfileLoaded())
            {
                Load += (s, e) => Close();
                return;
            }

            // Fix any illegal characters in the user name string
            strUserID = myStringFuncs.StripIllegals(Environment.UserName, "_", false);

            // The XML has loaded OK. Try to connect to the database and obtain the required info.
            myArcSDEFuncs = new ArcSDEFunctions();
            mySQLServerFuncs = new SQLServerFunctions();
            myFileFuncs = new FileFunctions();

            //// Check if the output folders exist.
            //if (!myFileFuncs.DirExists(myConfig.GetLogFilePath()))
            //{
            //    MessageBox.Show("The log file path " + myConfig.GetLogFilePath() + " does not exist. Cannot load Data Extractor.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    blOpenForm = false;
            //}

            ////if (!myFileFuncs.DirExists(myConfig.GetDefaultPath()) && blOpenForm)
            //{
            //    MessageBox.Show("The output path " + myConfig.GetDefaultPath() + " does not exist. Cannot load Data Extractor.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    blOpenForm = false;
            //}

            // Check if SDE file exists. If not, bail.
            string strSDE = myConfig.GetSDEName();
            if (!myFileFuncs.FileExists(strSDE) && blOpenForm)
            {
                MessageBox.Show("ArcSDE connection file " + strSDE + " not found. Cannot load Data Extractor.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                blOpenForm = false;
            }

            // Close the form if there are any errors at this point.
            if (!blOpenForm)
            {
                Load += (s, e) => Close();
                return;
            }

            // SDE file exists; let's try to open it.
            IWorkspace wsSQLWorkspace = null;
            if (blOpenForm)
            {
                try
                {
                    wsSQLWorkspace = myArcSDEFuncs.OpenArcSDEConnection(strSDE);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Cannot open ArcSDE connection " + strSDE + ". Error is " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    blOpenForm = false;
                }
            }

            // Get the list of SQL tables we can use.
            string strIncludeWildcard = myConfig.GetIncludeWildcard();
            string strExcludeWildcard = myConfig.GetExcludeWildcard();
            strTableList = myArcSDEFuncs.GetTableNames(wsSQLWorkspace, strIncludeWildcard, strExcludeWildcard);

            // Get all of the partner layer columns
            string strPartnerTable = myConfig.GetPartnerTable();
            string strPartnerColumn = myConfig.GetPartnerColumn();
            string strActiveColumn = myConfig.GetActiveColumn();
            string strShortColumn = myConfig.GetShortColumn();
            string strFormatColumn = myConfig.GetFormatColumn();
            string strExportColumn = myConfig.GetExportColumn();
            string strSQLTableColumn = myConfig.GetSQLTableColumn();
            string strNotesColumn = myConfig.GetNotesColumn();
            string strPartnerClause = myConfig.GetPartnerClause();

            // Make sure the table and columns exist.
            if (myArcMapFuncs.LayerExists(strPartnerTable))
            {
                // Check all of the partner columns are in the partner table.
                List<string> liAllPartnerColumns = myConfig.GetAllPartnerColumns();
                List<string> liExistingPartnerColumns = myArcMapFuncs.FieldsExist(strPartnerTable, liAllPartnerColumns);
                var theDifference = liAllPartnerColumns.Except(liExistingPartnerColumns).ToList();
                if (theDifference.Count != 0)
                {
                    string strMessage = "The column(s) ";
                    foreach (string strCol in theDifference)
                    {
                        strMessage = strMessage + strCol + ", ";
                    }
                    strMessage = strMessage.Substring(0, strMessage.Length - 2) + " could not be found in table " + strPartnerTable + ". Cannot load Data Extractor.";
                    MessageBox.Show(strMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    blOpenForm = false;
                }
            }
            else
            {
                MessageBox.Show("Partner table '" + strPartnerTable + "' not found. Cannot load Data Extractor.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                blOpenForm = false;
            }

            // Decide whether the form should open at all and if not return control to the application.
            if (!blOpenForm)
            {
                Load += (s, e) => Close();
                return;
            }

            // Set the default partner where clause
            string strWhereClause = strPartnerClause;
            if (strWhereClause == "")
                strWhereClause = strActiveColumn + " = 'Y'";

            // Load all the active partners to the list.
            ILayer pPartnerLayer = myArcMapFuncs.GetLayer(strPartnerTable);
            IFeatureLayer pPartnerFL;
            try
            {
                pPartnerFL = (IFeatureLayer)pPartnerLayer;
                IFeatureClass pPartnerFC = pPartnerFL.FeatureClass;
                IQueryFilter myFilter = new QueryFilter();
                myFilter.WhereClause = strWhereClause;
                IFeatureCursor pPartnerCurs = pPartnerFC.Search(myFilter, false);
                int intPartnerColumn = pPartnerFC.FindField(strPartnerColumn);
                int intShortColumn = pPartnerFC.FindField(strShortColumn);
                int intFormatColumn = pPartnerFC.FindField(strFormatColumn);
                int intExportColumn = pPartnerFC.FindField(strExportColumn);
                int intSQLTableColumn = pPartnerFC.FindField(strSQLTableColumn);
                int intNotesColumn = pPartnerFC.FindField(strNotesColumn);

                ITableSort tableSort = new TableSortClass();
                tableSort.Table = (ITable)pPartnerFC;
                tableSort.Cursor = (ICursor)pPartnerCurs;
                tableSort.Fields = strPartnerColumn;
                tableSort.set_Ascending(strPartnerColumn, true);
                tableSort.Sort(null);

                ICursor pSortCurs = tableSort.Rows;
                IRow pRow = null;
                List<string> strPartnerList = new List<string>();
                while ((pRow = pSortCurs.NextRow()) != null)
                {
                    // Add the partner name to the list box
                    lstActivePartners.Items.Add(pRow.get_Value(intPartnerColumn).ToString());

                    // Load all of the partner details to the lists
                    liActivePartnerNames.Add(pRow.get_Value(intPartnerColumn).ToString());
                    liActivePartnerShorts.Add(pRow.get_Value(intShortColumn).ToString());
                    liActivePartnerFormats.Add(pRow.get_Value(intFormatColumn).ToString());
                    liActivePartnerExports.Add(pRow.get_Value(intExportColumn).ToString());
                    liActivePartnerSQLTables.Add(pRow.get_Value(intSQLTableColumn).ToString());
                    liActivePartnerNotes.Add(pRow.get_Value(intNotesColumn).ToString());
                }
                pPartnerCurs = null;
                pSortCurs = null;
                myFilter = null;
                tableSort = null;
            }
            catch
            {
                MessageBox.Show("Layer " + strPartnerTable + " is not a feature layer.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                blOpenForm = false;
            }

            // Now load all the SQL layers that are listed in the XML profile
            // into the menu.
            liSQLFiles = myConfig.GetSQLTables(); //Node names as defined in the SQLFileList.
            liSQLOutputNames = myConfig.GetSQLOutputNames(); // Output names for each node.
            liSQLOutputTypes = myConfig.GetSQLOutputTypes(); // Output types for each node.
            liSQLColumns = myConfig.GetSQLColumns();
            liSQLWhereClauses = myConfig.GetSQLWhereClauses();
            liSQLOrderClauses = myConfig.GetSQLOrderClauses();
            liSQLMacroNames = myConfig.GetSQLMacroNames();
            liSQLMacroParms = myConfig.GetSQLMacroParms();

            int a = 0;
            foreach (string strFile in liSQLFiles)
            {
                string strItemEntry = strFile + " ---> " + liSQLOutputNames[a];
                lstTables.Items.Add(strItemEntry);

                liSQLTableEntries.Add(strItemEntry);
                a++;
            }

            // Now load all the map layers that are listed in the XML profile
            // into the menu, if they are loaded in the project.
            List<string> liAllLayers = myConfig.GetMapTableNames(); // The GIS layer names.
            List<string> liAllOutputNames = myConfig.GetMapOutputNames(); // The output names.
            List<string> liAllOutputTypes = myConfig.GetMapOutputTypes(); // The output types.
            List<string> liAllFiles = myConfig.GetMapTables(); // The node names.
            List<string> liAllColumns = myConfig.GetMapColumns();
            List<string> liAllWhereClauses = myConfig.GetMapWhereClauses();
            List<string> liAllOrderClauses = myConfig.GetMapOrderClauses();
            List<string> liAllMacroNames = myConfig.GetMapMacroNames();
            List<string> liAllMacroParms = myConfig.GetMapMacroParms();

            List<string> liMissingLayers = new List<string>();
            a = 0;
            foreach (string strLayer in liAllLayers) // For all possible source layers.
            {
                if (myArcMapFuncs.LayerExists(strLayer))
                {
                    string strItemEntry = liAllFiles[a] + " ---> " + strLayer;
                    lstLayers.Items.Add(strItemEntry);

                    // Add all the info for this item to the lists.
                    liOpenLayerEntries.Add(strItemEntry);
                    liOpenLayers.Add(strLayer);
                    liOpenOutputNames.Add(liAllOutputNames[a]);
                    liOpenOutputTypes.Add(liAllOutputTypes[a]);
                    liOpenFiles.Add(liAllFiles[a]);
                    liOpenColumns.Add(liAllColumns[a]);
                    liOpenWhereClauses.Add(liAllWhereClauses[a]);
                    liOpenOrderClauses.Add(liAllOrderClauses[a]);
                    liOpenMacroNames.Add(liAllMacroNames[a]);
                    liOpenMacroParms.Add(liAllMacroParms[a]);
                }
                else
                    liMissingLayers.Add(strLayer);
                a++;
            }

            // If there are no layers loaded
            if ((liMissingLayers.Count == liAllLayers.Count) && (liAllLayers.Count != 0))
            {
                MessageBox.Show("There are no GIS layers loaded in the view.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            // If there is one missing layer
            else if (liMissingLayers.Count == 1)
            {
                MessageBox.Show("The following GIS layer is not loaded: " + Environment.NewLine + Environment.NewLine +
                    liMissingLayers[0], "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            // If there are multiple missing layers
            else if (liMissingLayers.Count > 0)
            {
                string strMessage = "The following GIS layers are not loaded: " + Environment.NewLine + Environment.NewLine;
                foreach (string strMissingLayer in liMissingLayers)
                    strMessage = strMessage + strMissingLayer + ", ";
                strMessage = strMessage.Substring(0, strMessage.Length - 2);
                MessageBox.Show(strMessage, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            // Set the check boxes and dropdown lists.
            chkZip.Checked = myConfig.GetDefaultZip();
            chkApplyExclusion.Checked = myConfig.GetDefaultExclusion();
            chkClearLog.Checked = myConfig.GetDefaultClearLogFile();
            chkUseCentroids.Checked = myConfig.GetDefaultUseCentroids();
            chkUploadToServer.Checked = myConfig.GetDefaultUploadToServer();

            foreach (string anOption in myConfig.GetSelectTypeOptions())
                cmbSelectionType.Items.Add(anOption);

            if (myConfig.GetDefaultSelectType() != -1)
                cmbSelectionType.SelectedIndex = myConfig.GetDefaultSelectType();

            // Hide controls that were not requested.
            if (myConfig.GetExclusionClause() == "")
            {
                chkApplyExclusion.Hide();
            }
            else
            {
                chkApplyExclusion.Show();
            }

            if (myConfig.GetHideUseCentroids())
            {
                chkUseCentroids.Hide();
            }
            else
            {
                chkUseCentroids.Show();
            }

            if (myConfig.GetHideUploadToServer())
            {
                chkUploadToServer.Hide();
            }
            else
            {
                chkUploadToServer.Show();
            }

            // Tidy up.
            wsSQLWorkspace = null;
        
        }

        private bool XMLProfileLoaded()
        {
            // Get the application XML file name
            myLaunchConfig = new ExtractorToolLaunchConfig();
            myFileFuncs = new FileFunctions();

            // Check the application XML file is found and load it
            if (!myLaunchConfig.XMLFound)
            {
                MessageBox.Show("XML file 'DataExtractor.xml' not found; form cannot load.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (!myLaunchConfig.XMLLoaded)
            {
                MessageBox.Show("Error loading XML File 'DataExtractor.xml'; form cannot load.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            // Prompt the user for the XML profile name (or use the default name)
            string strXMLFolder = myFileFuncs.GetDirectoryName(Settings.Default.XMLFile);
            bool blOnlyDefault = true;
            int intCount = 0;
            if (myLaunchConfig.ChooseConfig) // If we are allowed to choose, check if there are multiple profiles. 
            // If there is only the default XML file in the directory, launch the form. Otherwise the user has to choose.
            {
                foreach (string strFileName in myFileFuncs.GetAllFilesInDirectory(strXMLFolder))
                {
                    if (myFileFuncs.GetFileName(strFileName).ToLower() != "dataextractor.xml" && myFileFuncs.GetExtension(strFileName).ToLower() == "xml")
                    {
                        // is it the default?
                        intCount++;
                        if (myFileFuncs.GetFileName(strFileName) != myLaunchConfig.DefaultXML)
                        {
                            blOnlyDefault = false;
                        }
                    }
                }
                if (intCount > 1)
                {
                    blOnlyDefault = false;
                }
            }
            if (myLaunchConfig.ChooseConfig && !blOnlyDefault)
            {
                // User has to choose the configuration file first.

                using (var myConfigForm = new frmChooseConfig(strXMLFolder, myLaunchConfig.DefaultXML))
                {
                    var result = myConfigForm.ShowDialog();
                    if (result == System.Windows.Forms.DialogResult.OK)
                    {
                        strConfigFile = strXMLFolder + @"\" + myConfigForm.ChosenXMLFile;
                    }
                    else
                    {
                        MessageBox.Show("No XML file was chosen; form cannot load.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }

            }
            else
            {
                strConfigFile = strXMLFolder + @"\" + myLaunchConfig.DefaultXML; // don't allow the user to choose, just use the default.
                // Just check it exists, though.
                if (!myFileFuncs.FileExists(strConfigFile))
                {
                    MessageBox.Show("The default XML file '" + myLaunchConfig.DefaultXML + "' was not found in the XML directory; form cannot load.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }

            // Load the required XML profile
            myConfig = new ExtractorToolConfig(strConfigFile); // Must now pass the correct XML name.

            IApplication pApp = ArcMap.Application;
            myArcMapFuncs = new ArcMapFunctions(pApp);
            myStringFuncs = new StringFunctions();

            // Did we find the XML?
            if (myConfig.GetFoundXML() == false)
            {
                MessageBox.Show("XML file not found; form cannot load.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            // Did it load OK?
            if (myConfig.GetLoadedXML() == false)
            {
                MessageBox.Show("Error loading XML File; form cannot load.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;

        }

        private void btnAbout_Click(object sender, EventArgs e)
        {
            // Remove focus from about button
            btnOK.Focus();

            // Show the about details
            string strTitle = "About Data Extractor";
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            string strText = "Data Extractor\r\n\r\n" +
                string.Format("Version {0}.{1}.{2}", version.Major, version.Minor, version.Build) + "\r\n\r\n" +
                "Copyright© 2017 SxBRC, 2017-2019 TVERC, 2020 Andy Foy Consulting.\r\n\r\n" +
                "Created by Hester Lyons Consulting & Andy Foy Consulting.\r\n\r\n" +
                "This program will extract SQL records and/or GIS layers\r\n" +
                "that intersect with partner boundary feature(s).\r\n" +
                "The user can select which SQL table or GIS layers to \r\n" +
                "extract and which partner boundaries to extract for.\r\n\r\n" +
                "This tool was developed with funding from:\r\n\r\n" +
                "  - Sussex Biodiversity Record Centre\r\n" +
                "  - Thames Valley Environmental Records Centre\r\n" +
                "  - Andy Foy Consulting";
            MessageBox.Show(strText, strTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            // User has clicked OK. 
            this.Cursor = Cursors.WaitCursor;

            // Firstly check all the information. If anything doesn't add up, bail out.

            // Have they selected at least one partner?
            if (lstActivePartners.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please select at least one partner.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.BringToFront();
                this.Cursor = Cursors.Default;
                return;
            }

            // Have they selected either a table or a GIS layer?
            if (lstTables.SelectedItems.Count == 0 && lstLayers.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please select one or more layers.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.BringToFront();
                this.Cursor = Cursors.Default;
                return;
            }

            // Have they selected an output type?
            if (cmbSelectionType.SelectedItem.ToString() == "")
            {
                MessageBox.Show("Please choose a selection type.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.BringToFront();
                this.Cursor = Cursors.Default;
                return;
            }

            // So far so good. Start the process.

            //================================================ PROCESS ===========================================================
            
            // Get the selected partners.
            List<string> liChosenPartners = new List<string>();
            foreach (string aPartner in lstActivePartners.SelectedItems)
            {
                liChosenPartners.Add(aPartner);
            }

            // Get the selected GIS layers
            List<string> liChosenSQLFiles = new List<string>();
            List<string> liChosenSQLOutputNames = new List<string>();
            List<string> liChosenSQLOutputTypes = new List<string>();
            List<string> liChosenSQLColumns = new List<string>();
            List<string> liChosenSQLWhereClauses = new List<string>();
            List<string> liChosenSQLOrderClauses = new List<string>();
            List<string> liChosenSQLMacroNames = new List<string>();
            List<string> liChosenSQLMacroParms = new List<string>();

            foreach (string aSQLFile in lstTables.SelectedItems)
            {
                // Find this entry to get the index.
                int a = liSQLTableEntries.IndexOf(aSQLFile);
                // Populate the other lists.
                liChosenSQLFiles.Add(liSQLFiles[a]);
                liChosenSQLOutputNames.Add(liSQLOutputNames[a]);
                liChosenSQLOutputTypes.Add(liSQLOutputTypes[a]);
                liChosenSQLColumns.Add(liSQLColumns[a]);
                liChosenSQLWhereClauses.Add(liSQLWhereClauses[a]);
                liChosenSQLOrderClauses.Add(liSQLOrderClauses[a]);
                liChosenSQLMacroNames.Add(liSQLMacroNames[a]);
                liChosenSQLMacroParms.Add(liSQLMacroParms[a]);
            }

            // Get the selected GIS layers
            List<string> liChosenGISLayers = new List<string>();
            List<string> liChosenGISOutputNames = new List<string>();
            List<string> liChosenGISOutputTypes = new List<string>();
            List<string> liChosenGISFiles = new List<string>();
            List<string> liChosenGISColumns = new List<string>();
            List<string> liChosenGISWhereClauses = new List<string>();
            List<string> liChosenGISOrderClauses = new List<string>();
            List<string> liChosenGISMacroNames = new List<string>();
            List<string> liChosenGISMacroParms = new List<string>();
            //List<string> liChosenGISLayerEntries = new List<string>();

            foreach (string aGISLayer in lstLayers.SelectedItems)
            {
                //liChosenGISLayerEntries.Add(aGISLayer);
                // Find this entry to get the index.
                int a = liOpenLayerEntries.IndexOf(aGISLayer);
                // Populate the other lists.
                liChosenGISLayers.Add(liOpenLayers[a]);
                liChosenGISOutputNames.Add(liOpenOutputNames[a]);
                liChosenGISOutputTypes.Add(liOpenOutputTypes[a]);
                liChosenGISFiles.Add(liOpenFiles[a]);
                liChosenGISColumns.Add(liOpenColumns[a]);
                liChosenGISWhereClauses.Add(liOpenWhereClauses[a]);
                liChosenGISOrderClauses.Add(liOpenOrderClauses[a]);
                liChosenGISMacroNames.Add(liOpenMacroNames[a]);
                liChosenGISMacroParms.Add(liOpenMacroParms[a]);
            }

            string strSelectionType = cmbSelectionType.SelectedItem.ToString();
            bool blCreateZip = chkZip.Checked;
            bool blApplyExclusion = chkApplyExclusion.Checked;
            bool blClearLogFile = chkClearLog.Checked;
            bool blUseCentroids = chkUseCentroids.Checked;
            bool blUploadToServer = chkUploadToServer.Checked;

            string strUseCentroids = "0"; // Use polygons is the default;
            if (blUseCentroids)
                strUseCentroids = "1";

            string strSelectionDigit = "1"; // Spatial is the default.
            // Check the selection type and set the appropriate number.
            if (strSelectionType.ToLower().Contains("spatial"))
            {
                if (strSelectionType.ToLower().Contains("tags"))
                {
                    strSelectionDigit = "3";
                }
            }
            else
            {
                strSelectionDigit = "2";
            }

            // Everything has been taken from the menu. Get some further data from the Config file.
            bool blDebugMode = myConfig.GetDebugMode();
            string strLogFilePath = myConfig.GetLogFilePath();
            string strDefaultPath = myConfig.GetDefaultPath();
            string strPartnerFolder = myConfig.GetPartnerFolder();
            string strSDEName = myConfig.GetSDEName();
            string strConnectionString = myConfig.GetConnectionString();
            string strDatabaseSchema = myConfig.GetDatabaseSchema();

            string strPartnerTable = myConfig.GetPartnerTable();
            string strPartnerColumn = myConfig.GetPartnerColumn();
            string strShortColumn = myConfig.GetShortColumn();
            string strNotesColumn = myConfig.GetNotesColumn();
            string strFormatColumn = myConfig.GetFormatColumn();
            string strExportColumn = myConfig.GetExportColumn();
            string strSQLTableColumn = myConfig.GetSQLTableColumn();
            string strSQLFilesColumn = myConfig.GetSQLFilesColumn();
            string strMapFilesColumn = myConfig.GetMapFilesColumn();
            string strTagsColumn = myConfig.GetTagsColumn();
            string strPartnerSpatialColumn = myConfig.GetSpatialColumn();

            string strExclusionClause = myConfig.GetExclusionClause();

            // Set the date variables
            DateTime dtNow = DateTime.Now;
            string strDateDD = dtNow.ToString("dd");
            string strDateMM = dtNow.ToString("MM");
            string strDateMMM = dtNow.ToString("MMM");
            string strDateMMMM = dtNow.ToString("MMMM");
            string strDateYY = dtNow.ToString("yy");
            double dQtr = (Math.Ceiling(dtNow.Month / 3.0 + 2) % 4) + 1;
            string strDateQQ = dQtr.ToString("00");
            string strDateYYYY = dtNow.ToString("yyyy");
            string strDateFFFF = myStringFuncs.GetFinancialYear(dtNow);

            // Replace any date variables in the log file path
            strLogFilePath = strLogFilePath.Replace("%dd%", strDateDD).Replace("%mm%", strDateMM).Replace("%mmm%", strDateMMM).Replace("%mmmm%", strDateMMMM);
            strLogFilePath = strLogFilePath.Replace("%yy%", strDateYY).Replace("%qq%", strDateQQ).Replace("%yyyy%", strDateYYYY).Replace("%ffff%", strDateFFFF);

            // Replace any date variables in the default path
            strDefaultPath = strDefaultPath.Replace("%dd%", strDateDD).Replace("%mm%", strDateMM).Replace("%mmm%", strDateMMM).Replace("%mmmm%", strDateMMMM);
            strDefaultPath = strDefaultPath.Replace("%yy%", strDateYY).Replace("%qq%", strDateQQ).Replace("%yyyy%", strDateYYYY).Replace("%ffff%", strDateFFFF);

            // Create the logfile path if it doesn't exist.
            if (!myFileFuncs.DirExists(strLogFilePath))
            {
                try
                {
                    Directory.CreateDirectory(strLogFilePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Cannot create directory " + strLogFilePath + ". System error: " + ex.Message);
                    this.BringToFront();
                    this.Cursor = Cursors.Default;
                    return;
                }
            }

            // Archive the log file (if it exists)
            string strLogFile = strLogFilePath + @"\DataExtractor_" + strUserID + ".log";
            if (blClearLogFile)
            {
                if (myFileFuncs.FileExists(strLogFile))
                {
                    // Get the last modified date/time for the current log file
                    DateTime dtLogFile = File.GetLastWriteTime(strLogFile);
                    string strLastMod = dtLogFile.ToString("yyyy") + dtLogFile.ToString("MM") + dtLogFile.ToString("dd") + "_" +
                        dtLogFile.ToString("HH") + dtLogFile.ToString("mm") + dtLogFile.ToString("ss");

                    // Rename the current log file
                    string strLogFileArchive = strLogFilePath + @"\DataExtractor_" + strUserID + "_" + strLastMod + ".log";
                    bool blRenamed = myFileFuncs.RenameFile(strLogFile, strLogFileArchive);
                    if (!blRenamed)
                    {
                        MessageBox.Show("Cannot rename log file. Please make sure it is not open in another window");
                        this.BringToFront();
                        this.Cursor = Cursors.Default;
                        return;
                    }
                }
            }

            // Start the log file.
            myFileFuncs.WriteLine(strLogFile, "----------------------------------------------------------------------");
            myFileFuncs.WriteLine(strLogFile, "Process started!");
            myFileFuncs.WriteLine(strLogFile, "----------------------------------------------------------------------");

            myArcMapFuncs.ToggleDrawing();
            myArcMapFuncs.ToggleTOC();
            this.BringToFront();

            if (blDebugMode) myFileFuncs.WriteLine(strLogFile, "Debug Mode is ON.");
            if (blDebugMode) myFileFuncs.WriteLine(strLogFile, "User ID = '" + strUserID + "'");

            blCreateZip = false; // For the moment this part of the functionality is disabled.

            if (blCreateZip == false)
                myFileFuncs.WriteLine(strLogFile, "Leaving extract files unzipped ...");

            if (strUseCentroids == "1")
                myFileFuncs.WriteLine(strLogFile, "Selecting features using centroids ...");
            else
                myFileFuncs.WriteLine(strLogFile, "Selecting features using boundaries ...");

            myFileFuncs.WriteLine(strLogFile, "Performing selection type '" + strSelectionDigit + "' ...");
            //myFileFuncs.WriteLine(strLogFile, "Processing table = '" + strChosenSQLFile + "' ...");

            // Upload the partner table if required (and at least one SQL output is selected).
            if (blUploadToServer && lstTables.SelectedItems.Count != 0)
            {
                myFileFuncs.WriteLine(strLogFile, "");
                myFileFuncs.WriteLine(strLogFile, "Uploading partner table to server ...");
                //(string InWorkspace, string InDatasetName, string OutWorkspace, string OutDatasetName, string SortOrder, bool Messages = false)
                bool blResult = myArcMapFuncs.CopyFeatures(strPartnerTable, strSDEName + @"\" + strDatabaseSchema + "." + strPartnerTable, "", blDebugMode); // Copies both to GDB and shapefile.
                if (!blResult)
                {
                    //MessageBox.Show("Error uploading partner table - process terminated.");
                    myFileFuncs.WriteLine(strLogFile, "Error uploading partner table - process terminated.");
                    this.BringToFront();
                    this.Cursor = Cursors.Default;
                    return;
                }
                myArcMapFuncs.RemoveDataset(strSDEName, strDatabaseSchema, strPartnerTable); // Temporary layer is removed.
                myFileFuncs.WriteLine(strLogFile, "Upload to server complete.");
            }

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
                lblStatus.Text = "Partner: " + strPartner;
                lblStatus.Visible = true;
                lblStatus.Refresh();

                // Variables that will be filled.
                string strShortName = null;
                string strNotes = null;
                string strFormat = null;
                string strExport = null;
                string strSQLTable = null;
                string strSQLFiles = "";
                string strMapFiles = "";
                string strTags = null;

                // Get all the information on this partner. Firstly select the correct polygon.
                //myFileFuncs.WriteLine(strLogFile, "Selecting partner boundary and retrieving information ...");
                string strFilterClause = strPartnerColumn + " = '" + strPartner + "'";
                bool blResult = myArcMapFuncs.SelectLayerByAttributes(strPartnerTable, strFilterClause, "NEW_SELECTION", blDebugMode);
                if (!blResult)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error selecting partner boundary and retrieving information");
                    this.BringToFront();
                    this.Cursor = Cursors.Default;
                    return;
                }

                // Get the associated cursor.
                ICursor pCurs = myArcMapFuncs.GetCursorOnFeatureLayer(strPartnerTable);

                // Extract the information and report any strangeness.
                if (pCurs == null)
                {
                    MessageBox.Show("Could not retrieve information for partner " + strPartner + ". Aborting.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    myFileFuncs.WriteLine(strLogFile, "Could not retrieve information for partner " + strPartner + ". Aborting.");
                    this.BringToFront();
                    this.Cursor = Cursors.Default;
                    return;
                }

                IRow pRow = null;
                int i = 0;
                while ((pRow = pCurs.NextRow()) != null)
                {
                    int a;
                    if (i == 0) // Only take the first entry.
                    {
                        a = pRow.Fields.FindField(strShortColumn);
                        strShortName = pRow.get_Value(a).ToString(); // Short name

                        a = pRow.Fields.FindField(strNotesColumn);
                        strNotes = pRow.get_Value(a).ToString(); // Notes

                        a = pRow.Fields.FindField(strFormatColumn);
                        strFormat = pRow.get_Value(a).ToString().ToUpper().Trim(); // Output format

                        a = pRow.Fields.FindField(strExportColumn);
                        strExport = pRow.get_Value(a).ToString().ToUpper().Trim(); // Export format

                        a = pRow.Fields.FindField(strSQLTableColumn);
                        strSQLTable = pRow.get_Value(a).ToString(); // SQL Table

                        a = pRow.Fields.FindField(strSQLFilesColumn);
                        strSQLFiles = pRow.get_Value(a).ToString(); //SQL files

                        a = pRow.Fields.FindField(strMapFilesColumn);
                        strMapFiles = pRow.get_Value(a).ToString(); // Map files

                        a = pRow.Fields.FindField(strTagsColumn);
                        strTags = pRow.get_Value(a).ToString(); // SQL tags
                    }
                    i++;
                }

                if (i > 1)
                {
                    MessageBox.Show("There are duplicate entries for partner " + strPartner + " in the partner table. Using the first entry.");
                    myFileFuncs.WriteLine(strLogFile, "There are duplicate entries for partner " + strPartner + " in the partner table. Using the first entry.");
                }

                myFileFuncs.WriteLine(strLogFile, "");
                myFileFuncs.WriteLine(strLogFile, "----------------------------------------------------------------------");
                myFileFuncs.WriteLine(strLogFile, "Processing partner '" + strPartner + " (" + strShortName + ")' ...");

                // Replace any date variables in the partner folder
                string strOutputFolder = strPartnerFolder;
                strOutputFolder = strOutputFolder.Replace("%dd%", strDateDD).Replace("%mm%", strDateMM).Replace("%mmm%", strDateMMM).Replace("%mmmm%", strDateMMMM);
                strOutputFolder = strOutputFolder.Replace("%yy%", strDateYY).Replace("%qq%", strDateQQ).Replace("%yyyy%", strDateYYYY).Replace("%ffff%", strDateFFFF);

                // Replace the partner shortname in the output name
                strOutputFolder = Regex.Replace(strOutputFolder, "%partner%", strShortName, RegexOptions.IgnoreCase);

                // Set up the output folder.
                string strOutFolder = strDefaultPath + @"\" + strOutputFolder;
                if (!myFileFuncs.DirExists(strOutFolder))
                {
                    myFileFuncs.WriteLine(strLogFile, "Creating output path '" + strOutFolder + "'.");
                    try
                    {
                        Directory.CreateDirectory(strOutFolder);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Cannot create directory " + strOutFolder + ". System error: " + ex.Message);
                        myFileFuncs.WriteLine(strLogFile, "Cannot create directory " + strOutFolder + ". System error: " + ex.Message);
                        myArcMapFuncs.ToggleDrawing();
                        myArcMapFuncs.ToggleTOC();
                        lblStatus.Text = "";
                        this.BringToFront();
                        this.Cursor = Cursors.Default;
                        return;
                    }
                }

                // Create the connection.
                SqlConnection dbConn = mySQLServerFuncs.CreateSQLConnection(myConfig.GetConnectionString());

                //------------------------------------------------------------------
                // Let's do the SQL layers.
                //------------------------------------------------------------------

                if (strSQLTable == "")
                {
                    myFileFuncs.WriteLine(strLogFile, "Skipping SQL outputs - table name is blank.");
                }
                else if (!strTableList.Contains(strSQLTable))
                {
                    myFileFuncs.WriteLine(strLogFile, "Skipping SQL outputs - table '" + strSQLTable + "' - not found.");
                }
                else
                {
                    // Note under current setup the CurrentSpatialTable never changes.
                    if (strSQLTable != "" && liChosenSQLFiles.Count() > 0)
                    {
                        //myFileFuncs.WriteLine(strLogFile, "Processing SQL layers for partner " + strPartner + ".");
                        //int b = 0;
                        //int intThisLayer = 1;

                        // Firstly do the spatial/ tags selection.
                        int intTimeoutSeconds = myConfig.GetTimeoutSeconds();
                        string strIntermediateTable = strSQLTable + "_" + strUserID; // The output table from the HLSppSelection stored procedure.

                        // Delete the original subset (in case it still exists).
                        SqlCommand myDeleteCommand = mySQLServerFuncs.CreateSQLCommand(ref dbConn, "HLClearSpatialSubset", CommandType.StoredProcedure);
                        mySQLServerFuncs.AddSQLParameter(ref myDeleteCommand, "Schema", strDatabaseSchema);
                        mySQLServerFuncs.AddSQLParameter(ref myDeleteCommand, "SpeciesTable", strSQLTable);
                        mySQLServerFuncs.AddSQLParameter(ref myDeleteCommand, "UserId", strUserID);
                        try
                        {
                            //myFileFuncs.WriteLine(strLogFile, "Opening SQL Connection");
                            dbConn.Open();
                            //myFileFuncs.WriteLine(strLogFile, "Executing stored procedure to delete spatial subselection.");
                            string strRowsAffect = myDeleteCommand.ExecuteNonQuery().ToString();
                            // Close the connection again.
                            myDeleteCommand = null;
                            dbConn.Close();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Could not execute stored procedure 'HLClearSpatialSubset'. System returned the following message: " +
                                ex.Message);
                            myFileFuncs.WriteLine(strLogFile, "Could not execute stored procedure HLClearSpatialSubset. System returned the following message: " +
                                ex.Message);
                            this.Cursor = Cursors.Default;
                            dbConn.Close();

                            myArcMapFuncs.ToggleDrawing();
                            myArcMapFuncs.ToggleTOC();
                            lblStatus.Text = "";
                            this.BringToFront();
                            this.Cursor = Cursors.Default;
                            return;
                        }

                        // make the spatial / tags selection.
                        string strStoredSpatialAndTagsProcedure = "AFSelectSppRecords";
                        SqlCommand mySpatialCommand = null;
                        if (intTimeoutSeconds == 0)
                        {
                            mySpatialCommand = mySQLServerFuncs.CreateSQLCommand(ref dbConn, strStoredSpatialAndTagsProcedure, CommandType.StoredProcedure); // Note pass connection by ref here.
                        }
                        else
                        {
                            mySpatialCommand = mySQLServerFuncs.CreateSQLCommand(ref dbConn, strStoredSpatialAndTagsProcedure, CommandType.StoredProcedure, intTimeoutSeconds);
                        }

                        mySQLServerFuncs.AddSQLParameter(ref mySpatialCommand, "Schema", strDatabaseSchema);
                        mySQLServerFuncs.AddSQLParameter(ref mySpatialCommand, "PartnerTable", strPartnerTable);
                        mySQLServerFuncs.AddSQLParameter(ref mySpatialCommand, "PartnerColumn", strShortColumn); // Used for selection
                        mySQLServerFuncs.AddSQLParameter(ref mySpatialCommand, "Partner", strShortName); // Used for selection
                        mySQLServerFuncs.AddSQLParameter(ref mySpatialCommand, "TagsColumn", strTagsColumn);
                        mySQLServerFuncs.AddSQLParameter(ref mySpatialCommand, "PartnerSpatialColumn", strPartnerSpatialColumn);
                        mySQLServerFuncs.AddSQLParameter(ref mySpatialCommand, "SelectType", strSelectionDigit);
                        mySQLServerFuncs.AddSQLParameter(ref mySpatialCommand, "SpeciesTable", strSQLTable);
                        mySQLServerFuncs.AddSQLParameter(ref mySpatialCommand, "UserId", strUserID);
                        mySQLServerFuncs.AddSQLParameter(ref mySpatialCommand, "UseCentroids", strUseCentroids);
                        //mySQLServerFuncs.AddSQLParameter(ref mySpatialCommand, "Split", strSplit);

                        // Execute stored procedure.
                        int intCount = 0;
                        bool blSuccess = true;
                        try
                        {
                            //myFileFuncs.WriteLine(strLogFile, "Opening SQL Connection");
                            dbConn.Open();
                            myFileFuncs.WriteLine(strLogFile, "Executing spatial selection from '" + strSQLTable + "' ...");
                            string strRowsAffect = mySpatialCommand.ExecuteNonQuery().ToString();
                            myFileFuncs.WriteLine(strLogFile, "Spatial selection complete.");

                            blSuccess = mySQLServerFuncs.TableHasRows(ref dbConn, strIntermediateTable);
                            if (blSuccess)
                            {
                                intCount = mySQLServerFuncs.CountRows(ref dbConn, strIntermediateTable);
                                myFileFuncs.WriteLine(strLogFile, string.Format("{0:n0}", intCount) + " records selected.");
                            }
                            else
                            {
                                myFileFuncs.WriteLine(strLogFile, "Procedure returned no records.");
                            }
                            myFileFuncs.WriteLine(strLogFile, "");

                            // Close the connection again.
                            mySpatialCommand = null;
                            dbConn.Close();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Could not execute stored procedure 'AFSelectSppRecords'. System returned the following message: " +
                                ex.Message);
                            myFileFuncs.WriteLine(strLogFile, "Could not execute stored procedure 'AFSelectSppRecords'. System returned the following message: " +
                                ex.Message);
                            this.Cursor = Cursors.Default;
                            dbConn.Close();
                            myArcMapFuncs.ToggleDrawing();
                            myArcMapFuncs.ToggleTOC();
                            lblStatus.Text = "";
                            this.BringToFront();
                            this.Cursor = Cursors.Default;
                            return;
                        }

                        intLayerIndex = 0;
                        foreach (string strSQLFile in liChosenSQLFiles) // Output files, not input tables.
                        {
                            intExtractCnt = intExtractCnt + 1;
                            myFileFuncs.WriteLine(strLogFile, "Starting process " + intExtractCnt + " of " + intExtractTot + " ...");

                            string strChosenFile = liChosenSQLFiles[intLayerIndex];
                            string strOutputTable = liChosenSQLOutputNames[intLayerIndex];
                            string strOutputType = liChosenSQLOutputTypes[intLayerIndex].ToUpper().Trim();
                            string strColumns = liChosenSQLColumns[intLayerIndex];
                            string strWhereClause = liChosenSQLWhereClauses[intLayerIndex];
                            string strOrderClause = liChosenSQLOrderClauses[intLayerIndex];
                            string strMacroName = liChosenSQLMacroNames[intLayerIndex];
                            string strMacroParm = liChosenSQLMacroParms[intLayerIndex];

                            // Check the partner requires something
                            if (((strExport == "") && (strFormat == "") && (strOutputType == "")) || (strSQLFiles == ""))
                            {
                                myFileFuncs.WriteLine(strLogFile, "Skipping output = '" + strSQLFile + "' - not required.");
                            }
                            else
                            {
                                // Does the partner want this layer?
                                //if (liSQLFiles.Contains(strSQLFile))
                                if (strSQLFiles.ToLower().Contains(strSQLFile.ToLower()) || strSQLFiles.ToLower() == "all")
                                {
                                    // They do, process. 
                                    lblStatus.Text = strPartner + ": Processing SQL table " + strSQLFile + ".";
                                    lblStatus.Refresh();

                                    bool blSpatial = false;

                                    // Replace any date variables in the output name
                                    strOutputTable = strOutputTable.Replace("%dd%", strDateDD).Replace("%mm%", strDateMM).Replace("%mmm%", strDateMMM).Replace("%mmmm%", strDateMMMM);
                                    strOutputTable = strOutputTable.Replace("%yy%", strDateYY).Replace("%qq%", strDateQQ).Replace("%yyyy%", strDateYYYY).Replace("%ffff%", strDateFFFF);

                                    // Replace the partner shortname in the output name
                                    strOutputTable = Regex.Replace(strOutputTable, "%partner%", strShortName, RegexOptions.IgnoreCase);

                                    // Build a list of all of the columns required
                                    List<string> liFields = new List<string>();
                                    List<string> liRawFields = strColumns.Split(',').ToList();
                                    foreach (string strField in liRawFields)
                                    {
                                        liFields.Add(strField.Trim());
                                    }

                                    myFileFuncs.WriteLine(strLogFile, "Processing output '" + strSQLFile + "' ...");

                                    // Set the output format
                                    string strOutputFormat = strFormat;
                                    if ((strOutputType == "GDB") || (strOutputType == "SHP") || (strOutputType == "DBF"))
                                    {
                                        myFileFuncs.WriteLine(strLogFile, "Overriding the output type with '" + strOutputType + "' ...");
                                        strOutputFormat = strOutputType;
                                    }

                                    myFileFuncs.WriteLine(strLogFile, "Extracting '" + strOutputTable + "' ...");

                                    string strSpatialColumn = "";
                                    string strSplit = "0"; // Do we need to split for polys / points?

                                    // Is there a geometry field in the data requested?
                                    string[] strGeometryFields = { "SP_GEOMETRY", "Shape" }; // Expand as required.
                                    foreach (string strField in strGeometryFields)
                                    {
                                        if (strColumns.ToLower().Contains(strField.ToLower()))
                                        {
                                            blSpatial = true;
                                            strSplit = "1"; // To be passed to stored procedure.
                                            strSpatialColumn = strField; // To be passed to stored procedure
                                        }
                                    }

                                    // if '*' is used then check for geometry field in the table.
                                    if (strColumns == "*")
                                    {
                                        string strCheckTable = strDatabaseSchema + "." + strSQLTable;
                                        dbConn.Open();
                                        foreach (string strField in strGeometryFields)
                                        {
                                            if (mySQLServerFuncs.FieldExists(ref dbConn, strCheckTable, strField))
                                            {
                                                blSpatial = true;
                                                strSpatialColumn = strField;
                                                strSplit = "1";
                                            }
                                        }
                                        dbConn.Close();
                                    }

                                    // Do the attribute query. This splits the output into points and polygons as relevant.
                                    // Set the temporary table names and the stored procedure names.
                                    string strStoredProcedure = "HLSelectSppSubset"; // Default for all data
                                    //string strPolyFC = strDatabaseSchema + "." + strIntermediateTable + "_poly"; // Change these.
                                    //string strPointFC = strDatabaseSchema + "." + strIntermediateTable + "_point";
                                    //string strTempTable = strDatabaseSchema + "." + strIntermediateTable + "_flat";
                                    string strPolyFC = strIntermediateTable + "_poly"; // Change these.
                                    string strPointFC = strIntermediateTable + "_point";
                                    string strTempTable = strIntermediateTable + "_flat";

                                    SqlCommand myCommand = null;

                                    if (intTimeoutSeconds == 0)
                                    {
                                        myCommand = mySQLServerFuncs.CreateSQLCommand(ref dbConn, strStoredProcedure, CommandType.StoredProcedure); // Note pass connection by ref here.
                                    }
                                    else
                                    {
                                        myCommand = mySQLServerFuncs.CreateSQLCommand(ref dbConn, strStoredProcedure, CommandType.StoredProcedure, intTimeoutSeconds);
                                    }

                                    // Add the exclusion clause if required.
                                    if (strWhereClause == "")
                                    {
                                        if (blApplyExclusion && strExclusionClause != "") // Exclusion clause is to be applied.
                                        {
                                            strWhereClause = strExclusionClause; // Note WHERE is already included in the SP and needs not be repeated.
                                        }
                                    }
                                    else
                                    {
                                        if (blApplyExclusion && strExclusionClause != "") // Exclusion clause is to be applied.
                                        {
                                            strWhereClause = "(" + strWhereClause + ") AND (" + strExclusionClause + ")"; // There is a where clause with the exclusion clause applied
                                        }
                                    }

                                    mySQLServerFuncs.AddSQLParameter(ref myCommand, "Schema", strDatabaseSchema);
                                    mySQLServerFuncs.AddSQLParameter(ref myCommand, "SpeciesTable", strIntermediateTable);
                                    mySQLServerFuncs.AddSQLParameter(ref myCommand, "SpatialColumn", strSpatialColumn);
                                    mySQLServerFuncs.AddSQLParameter(ref myCommand, "ColumnNames", strColumns);
                                    mySQLServerFuncs.AddSQLParameter(ref myCommand, "WhereClause", strWhereClause);
                                    mySQLServerFuncs.AddSQLParameter(ref myCommand, "GroupByClause", "");
                                    mySQLServerFuncs.AddSQLParameter(ref myCommand, "OrderByClause", strOrderClause);
                                    mySQLServerFuncs.AddSQLParameter(ref myCommand, "UserID", strUserID);
                                    mySQLServerFuncs.AddSQLParameter(ref myCommand, "Split", strSplit);

                                    if (blDebugMode) myFileFuncs.WriteLine(strLogFile, "Spatial column is " + strSpatialColumn);
                                    myFileFuncs.WriteLine(strLogFile, "Columns clause is ... " + strColumns);
                                    if (strWhereClause.Length > 0)
                                        myFileFuncs.WriteLine(strLogFile, "Where clause is ... " + strWhereClause.Replace("\r\n", " "));
                                    else
                                        myFileFuncs.WriteLine(strLogFile, "No where clause is specified.");

                                    if (strOrderClause.Length > 0)
                                        myFileFuncs.WriteLine(strLogFile, "Order by clause is ... " + strOrderClause.Replace("\r\n", " "));
                                    else
                                        myFileFuncs.WriteLine(strLogFile, "No order by clause is specified.");

                                    if (strSplit == "1")
                                        if (blDebugMode) myFileFuncs.WriteLine(strLogFile, "Data will be split into points and polygons.");

                                    // Open SQL connection to database and
                                    // Run the stored procedure.
                                    int intPolyCount = 0;
                                    int intPointCount = 0;
                                    int intRowLength = 0;
                                    try
                                    {
                                        //myFileFuncs.WriteLine(strLogFile, "Opening SQL Connection");
                                        dbConn.Open();

                                        //myFileFuncs.WriteLine(strLogFile, "Executing stored procedure to make subset selection.");
                                        string strRowsAffect = myCommand.ExecuteNonQuery().ToString();
                                        if (blSpatial)
                                        {
                                            blSuccess = mySQLServerFuncs.TableHasRows(ref dbConn, strPointFC);
                                            if (!blSuccess)
                                                blSuccess = mySQLServerFuncs.TableHasRows(ref dbConn, strPolyFC);
                                        }
                                        else
                                            blSuccess = mySQLServerFuncs.TableHasRows(ref dbConn, strTempTable);

                                        if (blSuccess && blSpatial)
                                        {
                                            intPolyCount = mySQLServerFuncs.CountRows(ref dbConn, strPolyFC);
                                            intPointCount = mySQLServerFuncs.CountRows(ref dbConn, strPointFC);
                                            intCount = intPolyCount + intPointCount;

                                            myFileFuncs.WriteLine(strLogFile, string.Format("{0:n0}", intPointCount) + " points to extract.");
                                            myFileFuncs.WriteLine(strLogFile, string.Format("{0:n0}", intPolyCount) + " polygons to extract.");

                                            // Calculate the total row length for the table.
                                            if (intPolyCount > 0)
                                                intRowLength = myArcMapFuncs.GetRowLength(strSDEName, strDatabaseSchema + "." + strPolyFC);
                                            else if (intPointCount > 0)
                                                intRowLength = intRowLength + myArcMapFuncs.GetRowLength(strSDEName, strDatabaseSchema + "." + strPointFC);

                                        }
                                        else if (blSuccess)
                                        {
                                            intCount = mySQLServerFuncs.CountRows(ref dbConn, strTempTable);
                                            myFileFuncs.WriteLine(strLogFile, string.Format("{0:n0}", intCount) + " records to extract.");

                                            // Calculate the total row length for the table.
                                            if (intCount > 0)
                                                intRowLength = myArcMapFuncs.GetRowLength(strSDEName, strDatabaseSchema + "." + strTempTable);
                                        }
                                        else
                                        {
                                            myFileFuncs.WriteLine(strLogFile, "No records to extract.");
                                        }

                                        //myFileFuncs.WriteLine(strLogFile, "Closing SQL Connection");
                                        dbConn.Close();
                                        myCommand = null;
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show("Could not execute stored procedure 'HLSelectSppSubset'. System returned the following message: " +
                                            ex.Message);
                                        myFileFuncs.WriteLine(strLogFile, "Could not execute stored procedure 'HLSelectSppSubset'. System returned the following message: " +
                                            ex.Message);
                                        this.Cursor = Cursors.Default;
                                        dbConn.Close();
                                        myArcMapFuncs.ToggleDrawing();
                                        myArcMapFuncs.ToggleTOC();
                                        this.BringToFront();
                                        this.Cursor = Cursors.Default;
                                        return;
                                    }

                                    if (blSuccess)
                                    {
                                        // Check if the maximum record length will be exceeded
                                        if (intRowLength > 4000)
                                        {
                                            myFileFuncs.WriteLine(strLogFile, "Record length exceeds maximum of 4,000 bytes - selection cancelled.");
                                        }
                                        else
                                            if (blDebugMode) myFileFuncs.WriteLine(strLogFile, string.Format("Total record length = {0} bytes", intRowLength));

                                        // Display the total data size
                                        int intDataSize;
                                        intDataSize = ((intRowLength * intCount) / 1024) + 1;
                                        string strDataSizeKb = string.Format("{0:n0}", intDataSize);
                                        string strDataSizeMb = string.Format("{0:n2}", (double)intDataSize / 1024);
                                        string strDataSizeGb = string.Format("{0:n2}", (double)intDataSize / (1024 * 1024));

                                        if (intDataSize > (1024 * 1024))
                                        {
                                            myFileFuncs.WriteLine(strLogFile, string.Format("Total data size = {0} Kb ({1} Gb).", strDataSizeKb, strDataSizeGb));
                                        }
                                        else
                                        {
                                            if (intDataSize > 1024)
                                                myFileFuncs.WriteLine(strLogFile, string.Format("Total data size = {0} Kb ({1} Mb).", strDataSizeKb, strDataSizeMb));
                                            else
                                                myFileFuncs.WriteLine(strLogFile, string.Format("Total data size = {0} Kb.", strDataSizeKb));
                                        }

                                        // Check if the total data size will be exceeded
                                        if (intDataSize > (2 * 1024 * 1024))
                                            myFileFuncs.WriteLine(strLogFile, "Total data size exceeds maximum of 2 Gb - selection cancelled.");

                                        // Report if the data is not spatial
                                        if (!blSpatial)
                                            myFileFuncs.WriteLine(strLogFile, "Data is not spatial.");
                                    }

                                    lblStatus.Text = strPartner + ": Writing output for " + strSQLFile + " to GIS format.";
                                    lblStatus.Refresh();
                                    this.BringToFront();

                                    // Now output to shape or table as appropriate.
                                    blResult = false;
                                    if (blSuccess)
                                    {
                                        // Check the output geodatabase exists
                                        string strOutPath = strOutFolder;
                                        if (strOutputFormat == "GDB")
                                        {
                                            strOutPath = strOutPath + "\\" + strOutputFolder + ".gdb";
                                            if (!myFileFuncs.DirExists(strOutPath))
                                            {
                                                myFileFuncs.WriteLine(strLogFile, "Creating output geodatabase '" + strOutPath + "' ...");
                                                myArcMapFuncs.CreateGeodatabase(strOutPath);
                                                myFileFuncs.WriteLine(strLogFile, "Output geodatabase created.");
                                            }
                                        }

                                        string strOutLayer = strOutPath + @"\" + strOutputTable; //strSQLFile; //NOTE not poly or point - for non-spatial output.
                                        string strOutLayerPoint = strOutLayer + "_point";
                                        string strOutLayerPoly = strOutLayer + "_poly";

                                        if ((strOutputFormat == "GDB") || (strOutputFormat == "SHP") || (strOutputFormat == "DBF"))
                                        {
                                            //if (blSpatial && strOutputFormat != "DBF")
                                            if (strOutputFormat != "DBF")
                                            {
                                                if (strOutputFormat == "SHP")
                                                    strOutLayerPoint = strOutLayerPoint + ".shp";

                                                if (intPointCount > 0)
                                                {
                                                    if (blDebugMode) myFileFuncs.WriteLine(strLogFile, "Outputing point selection to '" + strOutLayerPoint + "' ...");
                                                    //strPointFC = strSDEName + @"\" + strPointFC;
                                                    blResult = myArcMapFuncs.CopyFeatures(strSDEName + @"\" + strDatabaseSchema + "." + strPointFC, strOutLayerPoint, "", blDebugMode); // Copies both to GDB and shapefile.
                                                    if (!blResult)
                                                    {
                                                        //MessageBox.Show("Error outputing point output from SQL table");
                                                        myFileFuncs.WriteLine(strLogFile, "Error outputing " + strPointFC + " to " + strOutLayerPoint);
                                                        this.Cursor = Cursors.Default;
                                                        myArcMapFuncs.ToggleDrawing();
                                                        myArcMapFuncs.ToggleTOC();
                                                        lblStatus.Text = "";
                                                        this.BringToFront();
                                                        this.Cursor = Cursors.Default;
                                                        return;
                                                    }
                                                    myArcMapFuncs.RemoveLayer(strOutputTable + "_point"); // Temporary layer is removed.
                                                }

                                                if (strOutputFormat == "SHP")
                                                    strOutLayerPoly = strOutLayerPoly + ".shp";

                                                if (intPolyCount > 0)
                                                {
                                                    if (blDebugMode) myFileFuncs.WriteLine(strLogFile, "Outputing polygon selection to '" + strOutLayerPoly + "' ...");
                                                    //strPolyFC = strSDEName + @"\" + strPolyFC;
                                                    blResult = myArcMapFuncs.CopyFeatures(strSDEName + @"\" + strDatabaseSchema + "." + strPolyFC, strOutLayerPoly, "", blDebugMode); // Copies both to GDB and shapefile.
                                                    if (!blResult)
                                                    {
                                                        //MessageBox.Show("Error outputing polygon output from SQL table");
                                                        myFileFuncs.WriteLine(strLogFile, "Error outputing " + strPolyFC + " to " + strOutLayerPoly);
                                                        this.Cursor = Cursors.Default;
                                                        myArcMapFuncs.ToggleDrawing();
                                                        myArcMapFuncs.ToggleTOC();
                                                        lblStatus.Text = "";
                                                        this.BringToFront();
                                                        this.Cursor = Cursors.Default;
                                                        return;
                                                    }
                                                    myArcMapFuncs.RemoveLayer(strOutputTable + "_poly"); // Temporary layer is removed.
                                                }

                                            }
                                            else // The output table is non-spatial.
                                            {
                                                if (strOutputFormat == "SHP" || strOutputFormat == "DBF")
                                                    strOutLayer = strOutLayer + ".dbf";

                                                //string strInTable = strSDEName + @"\" + strTempTable;
                                                if (blDebugMode) myFileFuncs.WriteLine(strLogFile, "Outputing selection to table '" + strOutLayer + "' ...");
                                                blResult = myArcMapFuncs.CopyTable(strSDEName + @"\" + strDatabaseSchema + "." + strTempTable, strOutLayer, "", blDebugMode);
                                                if (!blResult)
                                                {
                                                    //MessageBox.Show("Error outputing output from SQL table");
                                                    myFileFuncs.WriteLine(strLogFile, "Error outputing " + strTempTable + " to " + strOutLayer);
                                                    this.Cursor = Cursors.Default;
                                                    myArcMapFuncs.ToggleDrawing();
                                                    myArcMapFuncs.ToggleTOC();
                                                    lblStatus.Text = "";
                                                    this.BringToFront();
                                                    this.Cursor = Cursors.Default;
                                                    return;
                                                }
                                                myArcMapFuncs.RemoveStandaloneTable(strOutputTable); // Temporary table is removed.
                                            }

                                            myFileFuncs.WriteLine(strLogFile, "Extract complete.");
                                        }

                                        // Set the export format
                                        string strExportFormat = strExport;
                                        if ((strOutputType == "CSV") || (strOutputType == "TXT"))
                                        {
                                            myFileFuncs.WriteLine(strLogFile, "Overriding the export type with '" + strOutputType + "' ...");
                                            strExportFormat = strOutputType;
                                        }

                                        // Now export to CSV if required.
                                        if (strExportFormat == "CSV")
                                        {
                                            lblStatus.Text = strPartner + ": Writing output for " + strSQLFile + " to CSV file.";
                                            lblStatus.Refresh();

                                            //if (blSpatial && strOutputFormat != "DBF")
                                            if (strOutputFormat != "DBF")
                                            {

                                                bool blAppend = false;
                                                string strOutputFile = strOutFolder + @"\" + strOutputTable + ".csv";
                                                if (intPointCount > 0)
                                                {
                                                    myFileFuncs.WriteLine(strLogFile, "Exporting points as a CSV file ...");
                                                    blResult = myArcMapFuncs.CopyToCSV(strOutLayerPoint, strOutputFile, liFields, true, blAppend, blDebugMode);
                                                    if (!blResult)
                                                    {
                                                        //MessageBox.Show("Error exporting output table to CSV file " + strOutputFile);
                                                        myFileFuncs.WriteLine(strLogFile, "Error exporting output table to CSV file " + strOutputFile);
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
                                                    myFileFuncs.WriteLine(strLogFile, "Exporting polygons as a CSV file ...");
                                                    blResult = myArcMapFuncs.CopyToCSV(strOutLayerPoly, strOutputFile, liFields, true, blAppend, blDebugMode);
                                                    if (!blResult)
                                                    {
                                                        //MessageBox.Show("Error appending output table to CSV file " + strOutputFile);
                                                        myFileFuncs.WriteLine(strLogFile, "Error appending output table to CSV file " + strOutputFile);
                                                        this.Cursor = Cursors.Default;
                                                        myArcMapFuncs.ToggleDrawing();
                                                        myArcMapFuncs.ToggleTOC();
                                                        lblStatus.Text = "";
                                                        this.BringToFront();
                                                        this.Cursor = Cursors.Default;
                                                        return;
                                                    }
                                                }

                                                myFileFuncs.WriteLine(strLogFile, "Export complete.");

                                                // Post-process the export file (if required)
                                                if (strMacroName != "")
                                                {
                                                    myFileFuncs.WriteLine(strLogFile, "Processing the export file ...");

                                                    Process scriptProc = new Process();
                                                    scriptProc.StartInfo.FileName = @"wscript";
                                                    //scriptProc.StartInfo.WorkingDirectory = @"c:\scripts\"; //<---very important 

                                                    scriptProc.StartInfo.Arguments = "//B //Nologo vbscript.vbs";
                                                    scriptProc.StartInfo.Arguments = string.Format("\"{0}\" \"{1}\" \"{2}\" \"{3}\" \"{4}\"",
                                                        strMacroName, strMacroParm, strOutFolder, strOutputTable + "." + strExportFormat, strLogFile);

                                                    scriptProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden; //prevent console window from popping up

                                                    try
                                                    {
                                                        scriptProc.Start();
                                                        scriptProc.WaitForExit();  // <-- Optional if you want program running until your script exit

                                                        int exitcode = scriptProc.ExitCode;
                                                        scriptProc.Close();

                                                        if (exitcode != 0)
                                                            MessageBox.Show("Error executing macro " + strMacroName + ". Exit code is " + exitcode, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        MessageBox.Show("Error executing macro " + strMacroName + ". Error message is " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                    }

                                                    myFileFuncs.WriteLine(strLogFile, "Processing complete.");
                                                }
                                            }
                                            else
                                            {
                                                string strOutputFile = strOutFolder + @"\" + strOutputTable + ".csv";
                                                myFileFuncs.WriteLine(strLogFile, "Exporting as a CSV file ...");
                                                blResult = myArcMapFuncs.CopyToCSV(strOutLayer, strOutputFile, liFields, false, false, blDebugMode);
                                                if (!blResult)
                                                {
                                                    //MessageBox.Show("Error exporting output table to CSV file " + strOutputFile);
                                                    myFileFuncs.WriteLine(strLogFile, "Error exporting output table to CSV file " + strOutputFile);
                                                    this.Cursor = Cursors.Default;
                                                    myArcMapFuncs.ToggleDrawing();
                                                    myArcMapFuncs.ToggleTOC();
                                                    lblStatus.Text = "";
                                                    this.BringToFront();
                                                    this.Cursor = Cursors.Default;
                                                    return;
                                                }

                                                myFileFuncs.WriteLine(strLogFile, "Export complete.");

                                                // Post-process the export file (if required)
                                                if (strMacroName != "")
                                                {
                                                    myFileFuncs.WriteLine(strLogFile, "Processing the export file ...");

                                                    Process scriptProc = new Process();
                                                    scriptProc.StartInfo.FileName = @"wscript";
                                                    //scriptProc.StartInfo.WorkingDirectory = @"c:\scripts\"; //<---very important 

                                                    scriptProc.StartInfo.Arguments = "//B //Nologo vbscript.vbs";
                                                    scriptProc.StartInfo.Arguments = string.Format("\"{0}\" \"{1}\" \"{2}\" \"{3}\" \"{4}\"",
                                                        strMacroName, strMacroParm, strOutFolder, strOutputTable + "." + strExportFormat, strLogFile);

                                                    scriptProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden; //prevent console window from popping up

                                                    try
                                                    {
                                                        scriptProc.Start();
                                                        scriptProc.WaitForExit();  // <-- Optional if you want program running until your script exit

                                                        int exitcode = scriptProc.ExitCode;
                                                        scriptProc.Close();

                                                        if (exitcode != 0)
                                                            MessageBox.Show("Error executing macro " + strMacroName + ". Exit code is " + exitcode, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        MessageBox.Show("Error executing macro " + strMacroName + ". Error message is " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                    }

                                                    myFileFuncs.WriteLine(strLogFile, "Processing complete.");
                                                }
                                            }
                                        }
                                        else if (strExportFormat == "TXT")
                                        {
                                            lblStatus.Text = strPartner + ": Writing output for " + strSQLFile + " to TXT file.";
                                            lblStatus.Refresh();

                                            //if (blSpatial && strOutputFormat != "DBF")
                                            if (strOutputFormat != "DBF")
                                            {

                                                bool blAppend = false;
                                                string strOutputFile = strOutFolder + @"\" + strOutputTable + ".txt";
                                                if (intPointCount > 0)
                                                {
                                                    myFileFuncs.WriteLine(strLogFile, "Exporting points as a TXT file ...");
                                                    blResult = myArcMapFuncs.CopyToTabDelimitedFile(strOutLayerPoint, strOutputFile, liFields, true, blAppend, blDebugMode);
                                                    if (!blResult)
                                                    {
                                                        //MessageBox.Show("Error exporting output table to txt file " + strOutputFile);
                                                        myFileFuncs.WriteLine(strLogFile, "Error exporting output table to txt file " + strOutputFile);
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
                                                    myFileFuncs.WriteLine(strLogFile, "Exporting polygons as a TXT file ...");
                                                    blResult = myArcMapFuncs.CopyToTabDelimitedFile(strOutLayerPoly, strOutputFile, liFields, true, blAppend, blDebugMode);
                                                    if (!blResult)
                                                    {
                                                        //MessageBox.Show("Error appending output table to txt file " + strOutputFile);
                                                        myFileFuncs.WriteLine(strLogFile, "Error appending output table to txt file " + strOutputFile);
                                                        this.Cursor = Cursors.Default;
                                                        myArcMapFuncs.ToggleDrawing();
                                                        myArcMapFuncs.ToggleTOC();
                                                        lblStatus.Text = "";
                                                        this.BringToFront();
                                                        this.Cursor = Cursors.Default;
                                                        return;
                                                    }
                                                }
                                                myFileFuncs.WriteLine(strLogFile, "Export complete.");
                                            }
                                            else
                                            {
                                                string strOutputFile = strOutFolder + @"\" + strOutputTable + ".txt";
                                                myFileFuncs.WriteLine(strLogFile, "Exporting as a TXT file ...");
                                                blResult = myArcMapFuncs.CopyToTabDelimitedFile(strOutLayer, strOutputFile, liFields, false, false, blDebugMode);
                                                if (!blResult)
                                                {
                                                    //MessageBox.Show("Error exporting output table to txt file " + strOutputFile);
                                                    myFileFuncs.WriteLine(strLogFile, "Error exporting output table to txt file " + strOutputFile);
                                                    this.Cursor = Cursors.Default;
                                                    myArcMapFuncs.ToggleDrawing();
                                                    myArcMapFuncs.ToggleTOC();
                                                    lblStatus.Text = "";
                                                    this.BringToFront();
                                                    this.Cursor = Cursors.Default;
                                                    return;
                                                }
                                                myFileFuncs.WriteLine(strLogFile, "Export complete.");
                                            }
                                        }

                                        // Delete the output table if not required
                                        if ((strOutputFormat != strOutputType) && (strOutputType != ""))
                                        {
                                            if (blDebugMode) myFileFuncs.WriteLine(strLogFile, "Deleting the output table ...");
                                            if (blSpatial)
                                            {
                                                if (intPointCount > 0)
                                                {
                                                    blResult = myArcMapFuncs.DeleteTable(strOutLayerPoint, blDebugMode);
                                                    if (!blResult)
                                                    {
                                                        myFileFuncs.WriteLine(strLogFile, "Error deleting the output table");
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
                                                    blResult = myArcMapFuncs.DeleteTable(strOutLayerPoly, blDebugMode);
                                                    if (!blResult)
                                                    {
                                                        myFileFuncs.WriteLine(strLogFile, "Error deleting the output table");
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
                                                blResult = myArcMapFuncs.DeleteTable(strOutLayer, blDebugMode);
                                                if (!blResult)
                                                {
                                                    myFileFuncs.WriteLine(strLogFile, "Error deleting the output table");
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

                                    lblStatus.Text = strPartner + ": Deleting temporary subset tables";
                                    this.BringToFront();

                                    // Delete the temporary tables in the SQL database
                                    strStoredProcedure = "HLClearSppSubset";
                                    SqlCommand myCommand2 = mySQLServerFuncs.CreateSQLCommand(ref dbConn, strStoredProcedure, CommandType.StoredProcedure); // Note pass connection by ref here.
                                    mySQLServerFuncs.AddSQLParameter(ref myCommand2, "Schema", strDatabaseSchema);
                                    mySQLServerFuncs.AddSQLParameter(ref myCommand2, "SpeciesTable", strSQLTable);
                                    mySQLServerFuncs.AddSQLParameter(ref myCommand2, "UserId", strUserID);
                                    try
                                    {
                                        //myFileFuncs.WriteLine(strLogFile, "Opening SQL connection");
                                        dbConn.Open();
                                        //myFileFuncs.WriteLine(strLogFile, "Deleting temporary tables.");
                                        string strRowsAffect = myCommand2.ExecuteNonQuery().ToString();
                                        //myFileFuncs.WriteLine(strLogFile, "Closing SQL connection");
                                        dbConn.Close();
                                        myCommand2 = null;
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show("Could not execute stored procedure 'HLClearSppSubset'. System returned the following message: " +
                                            ex.Message);
                                        myFileFuncs.WriteLine(strLogFile, "Could not execute stored procedure 'HLClearSppSubset'. System returned the following message: " +
                                            ex.Message);
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
                                    myFileFuncs.WriteLine(strLogFile, "Skipping output = '" + strSQLFile + "' - not required.");
                                }
                            }

                            myFileFuncs.WriteLine(strLogFile, "Completed process " + intExtractCnt + " of " + intExtractTot + ".");
                            myFileFuncs.WriteLine(strLogFile, "");

                            intLayerIndex++;
                        }

                        // Delete the final temporary spatial table.
                        string strSP = "HLClearSpatialSubset";
                        SqlCommand myCommand3 = mySQLServerFuncs.CreateSQLCommand(ref dbConn, strSP, CommandType.StoredProcedure); // Note pass connection by ref here.
                        mySQLServerFuncs.AddSQLParameter(ref myCommand3, "Schema", strDatabaseSchema);
                        mySQLServerFuncs.AddSQLParameter(ref myCommand3, "SpeciesTable", strSQLTable);
                        mySQLServerFuncs.AddSQLParameter(ref myCommand3, "UserId", strUserID);
                        try
                        {
                            //myFileFuncs.WriteLine(strLogFile, "Opening SQL connection");
                            dbConn.Open();
                            //myFileFuncs.WriteLine(strLogFile, "Deleting temporary spatial tables.");
                            string strRowsAffect = myCommand3.ExecuteNonQuery().ToString();
                            //myFileFuncs.WriteLine(strLogFile, "Closing SQL connection");
                            dbConn.Close();
                            myCommand3 = null;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Could not execute stored procedure 'HLClearSpatialSubset'. System returned the following message: " +
                                ex.Message);
                            myFileFuncs.WriteLine(strLogFile, "Could not execute stored procedure 'HLClearSpatialSubset'. System returned the following message: " +
                                ex.Message);
                            myArcMapFuncs.ToggleDrawing();
                            myArcMapFuncs.ToggleTOC();
                            lblStatus.Text = "";
                            this.BringToFront();
                            this.Cursor = Cursors.Default;
                            return;
                        }
                    }
                }

                lblStatus.Text = "";
                lblStatus.Refresh();
                this.BringToFront();

                //------------------------------------------------------------------
                // Let's do the GIS layers.
                //------------------------------------------------------------------

                //myFileFuncs.WriteLine(strLogFile, "Processing GIS layers for partner " + strPartner + ".");

                intLayerIndex = 0;
                foreach (string strGISLayer in liChosenGISLayers)
                {
                    intExtractCnt = intExtractCnt + 1;
                    myFileFuncs.WriteLine(strLogFile, "Starting process " + intExtractCnt + " of " + intExtractTot + " ...");

                    string strChosenFile = liChosenGISFiles[intLayerIndex];

                    string strTableName = strGISLayer;
                    string strOutputName = liChosenGISOutputNames[intLayerIndex];
                    string strOutputType = liChosenGISOutputTypes[intLayerIndex].ToUpper().Trim();
                    string strLayerColumns = liChosenGISColumns[intLayerIndex];
                    string strLayerWhereClause = liChosenGISWhereClauses[intLayerIndex];
                    string strLayerOrderClause = liChosenGISOrderClauses[intLayerIndex];
                    string strMacroName = liChosenGISMacroNames[intLayerIndex];
                    string strMacroParm = liChosenGISMacroParms[intLayerIndex];

                    // Check the partner requires something
                    if (((strExport == "") && (strFormat == "") && (strOutputType == "")) || (strMapFiles == ""))
                    {
                        myFileFuncs.WriteLine(strLogFile, "Skipping output = '" + strGISLayer + "' - not required.");
                    }
                    else
                    {
                        // Does the partner want this layer?
                        //if (liMapFiles.Contains(strChosenFile))
                        if (strMapFiles.ToLower().Contains(strChosenFile.ToLower()) || strMapFiles.ToLower() == "all")
                        {
                            lblStatus.Text = strPartner + ": Processing GIS layer " + strChosenFile + ".";
                            lblStatus.Refresh();
                            this.BringToFront();

                            // Replace any date variables in the output name
                            strOutputName = strOutputName.Replace("%dd%", strDateDD).Replace("%mm%", strDateMM).Replace("%mmm%", strDateMMM).Replace("%mmmm%", strDateMMMM);
                            strOutputName = strOutputName.Replace("%yy%", strDateYY).Replace("%qq%", strDateQQ).Replace("%yyyy%", strDateYYYY).Replace("%ffff%", strDateFFFF);

                            // Replace the partner shortname in the output name
                            strOutputName = Regex.Replace(strOutputName, "%partner%", strShortName, RegexOptions.IgnoreCase);

                            // Build a list of all of the columns required
                            List<string> liFields = new List<string>();
                            List<string> liRawFields = strLayerColumns.Split(',').ToList();
                            foreach (string strField in liRawFields)
                            {
                                liFields.Add(strField.Trim());
                            }

                            myFileFuncs.WriteLine(strLogFile, "Processing output '" + strChosenFile + "' ...");

                            // Set the output format
                            string strOutputFormat = strFormat;
                            if ((strOutputType == "GDB") || (strOutputType == "SHP") || (strOutputType == "DBF"))
                            {
                                myFileFuncs.WriteLine(strLogFile, "Overriding the output type with '" + strOutputType + "' ...");
                                strOutputFormat = strOutputType;
                            }

                            myFileFuncs.WriteLine(strLogFile, "Executing spatial selection ...");

                            // if '*' is used then it must be spatial.
                            bool blSpatial = false;
                            if (strLayerColumns == "*")
                            {
                                blSpatial = true;
                            }
                            else
                            {
                                // Is there a geometry field in the data requested?
                                string[] strGeometryFields = { "SP_GEOMETRY", "Shape" }; // Expand as required.
                                foreach (string strField in strGeometryFields)
                                {
                                    if (strLayerColumns.ToLower().Contains(strField.ToLower()))
                                    {
                                        blSpatial = true;
                                    }
                                }
                            }

                            // Firstly do the spatial selection.
                            blResult = myArcMapFuncs.SelectLayerByLocation(strTableName, strPartnerTable, "INTERSECT", "", "NEW_SELECTION", blDebugMode);
                            if (!blResult)
                            {
                                myFileFuncs.WriteLine(strLogFile, "Error creating selection using spatial query");
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
                                if (blDebugMode) myFileFuncs.WriteLine(strLogFile, "Creating subset selection for layer " + strTableName + " using attribute query:");

                                myFileFuncs.WriteLine(strLogFile, strLayerWhereClause);
                                blResult = myArcMapFuncs.SelectLayerByAttributes(strTableName, strLayerWhereClause, "SUBSET_SELECTION", blDebugMode);
                                if (!blResult)
                                {
                                    myFileFuncs.WriteLine(strLogFile, "Error creating subset selection using attribute query");
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

                            myFileFuncs.WriteLine(strLogFile, "Columns clause is ... " + strLayerColumns);
                            if (strLayerWhereClause.Length > 0)
                                myFileFuncs.WriteLine(strLogFile, "Where clause is ... " + strLayerWhereClause.Replace("\r\n", " "));
                            else
                                myFileFuncs.WriteLine(strLogFile, "No where clause is specified.");
                            if (strLayerOrderClause.Length > 0)
                                myFileFuncs.WriteLine(strLogFile, "Order by clause is ... " + strLayerOrderClause.Replace("\r\n", " "));
                            else
                                myFileFuncs.WriteLine(strLogFile, "No order by clause is specified.");

                            // If there is a selection, process it.
                            if (intSelectedFeatures > 0)
                            {
                                myFileFuncs.WriteLine(strLogFile, string.Format("{0:n0}", intSelectedFeatures) + " records selected.");

                                lblStatus.Text = strPartner + ": Exporting selection in layer " + strChosenFile + ".";
                                lblStatus.Refresh();
                                this.BringToFront();

                                // Export to Shapefile. Create new folder if required.

                                // Check the output geodatabase exists
                                string strOutPath = strOutFolder;
                                if (strOutputFormat == "GDB")
                                {
                                    strOutPath = strOutPath + "\\" + strOutputFolder + ".gdb";
                                    if (!myFileFuncs.DirExists(strOutPath))
                                    {
                                        myFileFuncs.WriteLine(strLogFile, "Creating output geodatabase '" + strOutPath + "' ...");
                                        myArcMapFuncs.CreateGeodatabase(strOutPath);
                                        myFileFuncs.WriteLine(strLogFile, "Output geodatabase created.");
                                    }
                                }

                                string strOutLayer = strOutPath + @"\" + strOutputName;

                                // Now export to shape or table as appropriate.
                                blResult = false;
                                if ((strOutputFormat == "GDB") || (strOutputFormat == "SHP") || (strOutputFormat == "DBF"))
                                {
                                    //if (blSpatial && strOutputFormat != "DBF")
                                    if (strOutputFormat != "DBF")
                                    {
                                        if (strOutputFormat == "SHP")
                                            strOutLayer = strOutLayer + ".shp";

                                        if (blDebugMode) myFileFuncs.WriteLine(strLogFile, "Exporting selection to GIS file: " + strOutLayer + ".");
                                        blResult = myArcMapFuncs.CopyFeatures(strTableName, strOutLayer, strLayerOrderClause, blDebugMode); // Copies both to GDB and shapefile.
                                        if (!blResult)
                                        {
                                            //MessageBox.Show("Error exporting output from " + strTableName);
                                            myFileFuncs.WriteLine(strLogFile, "Error exporting " + strTableName + " to " + strOutLayer);
                                            this.Cursor = Cursors.Default;
                                            myArcMapFuncs.ToggleDrawing();
                                            myArcMapFuncs.ToggleTOC();
                                            lblStatus.Text = "";
                                            this.BringToFront();
                                            this.Cursor = Cursors.Default;
                                            return;
                                        }

                                        //if (strOutputFormat == "SHP") // Not needed as ArcGIS does it automatically???
                                        //{
                                        //    if (blDebugMode) myFileFuncs.WriteLine(strLogFile, "Add spatial index to GIS file: " + strOutLayer + ".");
                                        //    blResult = myArcMapFuncs.AddSpatialIndex(strTableName);
                                        //    if (!blResult)
                                        //    {
                                        //        MessageBox.Show("Error adding spatial index to " + strOutLayer);
                                        //        myFileFuncs.WriteLine(strLogFile, "Error adding spatial index to " + strOutLayer);
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
                                        if (blDebugMode) myFileFuncs.WriteLine(strLogFile, "Removing non-required fields from the output file.");
                                        myArcMapFuncs.KeepSelectedFields(strOutLayer, liFields);
                                        myArcMapFuncs.RemoveLayer(strOutputName);
                                    }
                                    else // The output table is non-spatial.
                                    {
                                        if (strOutputFormat == "SHP" || strOutputFormat == "DBF")
                                            strOutLayer = strOutLayer + ".dbf";

                                        if (blDebugMode) myFileFuncs.WriteLine(strLogFile, "Exporting selection to table '" + strOutLayer + "' ...");
                                        blResult = myArcMapFuncs.CopyTable(strTableName, strOutLayer, strLayerOrderClause, blDebugMode);
                                        if (!blResult)
                                        {
                                            //MessageBox.Show("Error exporting output from " + strTableName);
                                            myFileFuncs.WriteLine(strLogFile, "Error exporting " + strTableName + " to " + strOutLayer);
                                            this.Cursor = Cursors.Default;
                                            myArcMapFuncs.ToggleDrawing();
                                            myArcMapFuncs.ToggleTOC();
                                            lblStatus.Text = "";
                                            this.BringToFront();
                                            this.Cursor = Cursors.Default;
                                            return;
                                        }

                                        // Drop non-required fields.
                                        if (blDebugMode) myFileFuncs.WriteLine(strLogFile, "Removing non-required fields from the output file.");
                                        myArcMapFuncs.KeepSelectedTableFields(strOutLayer, liFields);
                                        myArcMapFuncs.RemoveStandaloneTable(strOutputName);
                                    }

                                    myFileFuncs.WriteLine(strLogFile, "Extract complete.");
                                }

                                //------------------------------------------------------------------
                                // Export to text file if requested as well.
                                //------------------------------------------------------------------
                                lblStatus.Text = strPartner + ": Exporting layer " + strChosenFile + " to text.";
                                lblStatus.Refresh();
                                this.BringToFront();

                                // Set the export format
                                string strExportFormat = strExport;
                                if ((strOutputType == "CSV") || (strOutputType == "TXT"))
                                {
                                    myFileFuncs.WriteLine(strLogFile, "Overriding the export type with '" + strOutputType + "' ...");
                                    strExportFormat = strOutputType;
                                }

                                // Export to a CSV file
                                if (strExportFormat == "CSV")
                                {
                                    string strOutTable = strOutFolder + @"\" + strOutputName + ".csv";
                                    myFileFuncs.WriteLine(strLogFile, "Exporting as a CSV file ...");

                                    //if (blSpatial && strOutputFormat != "DBF")
                                    if (strOutputFormat != "DBF")
                                    {
                                        blResult = myArcMapFuncs.CopyToCSV(strOutLayer, strOutTable, liFields, true, false, blDebugMode);
                                        if (!blResult)
                                        {
                                            //MessageBox.Show("Error exporting output table to CSV file " + strOutputFile);
                                            myFileFuncs.WriteLine(strLogFile, "Error exporting output table to CSV file " + strOutTable);
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
                                        blResult = myArcMapFuncs.CopyToCSV(strOutLayer, strOutTable, liFields, false, false, blDebugMode);
                                        if (!blResult)
                                        {
                                            //MessageBox.Show("Error exporting output table to CSV file " + strOutputFile);
                                            myFileFuncs.WriteLine(strLogFile, "Error exporting output table to CSV file " + strOutTable);
                                            this.Cursor = Cursors.Default;
                                            myArcMapFuncs.ToggleDrawing();
                                            myArcMapFuncs.ToggleTOC();
                                            lblStatus.Text = "";
                                            this.BringToFront();
                                            this.Cursor = Cursors.Default;
                                            return;
                                        }
                                    }

                                    myFileFuncs.WriteLine(strLogFile, "Export complete.");

                                    // Post-process the export file (if required)
                                    if (strMacroName != "")
                                    {
                                        myFileFuncs.WriteLine(strLogFile, "Processing the export file ...");

                                        Process scriptProc = new Process();
                                        scriptProc.StartInfo.FileName = @"wscript";
                                        //scriptProc.StartInfo.WorkingDirectory = @"c:\scripts\"; //<---very important 

                                        scriptProc.StartInfo.Arguments = "//B //Nologo vbscript.vbs";
                                        scriptProc.StartInfo.Arguments = string.Format("\"{0}\" \"{1}\" \"{2}\" \"{3}\" \"{4}\"",
                                            strMacroName, strMacroParm, strOutFolder, strOutputName + "." + strExportFormat, strLogFile);

                                        scriptProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden; //prevent console window from popping up

                                        try
                                        {
                                            scriptProc.Start();
                                            scriptProc.WaitForExit();  // <-- Optional if you want program running until your script exit

                                            int exitcode = scriptProc.ExitCode;
                                            scriptProc.Close();

                                            if (exitcode != 0)
                                                MessageBox.Show("Error executing macro " + strMacroName + ". Exit code is " + exitcode, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show("Error executing macro " + strMacroName + ". Error message is " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }

                                        myFileFuncs.WriteLine(strLogFile, "Processing complete.");
                                    }
                                }
                                // Export to a TXT file
                                else if (strExportFormat == "TXT")
                                {
                                    string strOutTable = strOutFolder + @"\" + strOutputName + ".txt";
                                    myFileFuncs.WriteLine(strLogFile, "Exporting as a TXT file ...");

                                    //if (blSpatial && strOutputFormat != "DBF")
                                    if (strOutputFormat != "DBF")
                                    {
                                        blResult = myArcMapFuncs.CopyToTabDelimitedFile(strOutLayer, strOutTable, liFields, true, false, blDebugMode);
                                        if (!blResult)
                                        {
                                            //MessageBox.Show("Error exporting output table to txt file " + strOutputFile);
                                            myFileFuncs.WriteLine(strLogFile, "Error exporting output table to txt file " + strOutTable);
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
                                        blResult = myArcMapFuncs.CopyToTabDelimitedFile(strOutLayer, strOutTable, liFields, false, false, blDebugMode);
                                        if (!blResult)
                                        {
                                            //MessageBox.Show("Error exporting output table to txt file " + strOutputFile);
                                            myFileFuncs.WriteLine(strLogFile, "Error exporting output table to txt file " + strOutTable);
                                            this.Cursor = Cursors.Default;
                                            myArcMapFuncs.ToggleDrawing();
                                            myArcMapFuncs.ToggleTOC();
                                            lblStatus.Text = "";
                                            this.BringToFront();
                                            this.Cursor = Cursors.Default;
                                            return;
                                        }
                                    }

                                    myFileFuncs.WriteLine(strLogFile, "Export complete.");

                                    // Post-process the export file (if required)
                                    if (strMacroName != "")
                                    {
                                        myFileFuncs.WriteLine(strLogFile, "Processing the export file ...");

                                        Process scriptProc = new Process();
                                        scriptProc.StartInfo.FileName = @"wscript";
                                        //scriptProc.StartInfo.WorkingDirectory = @"c:\scripts\"; //<---very important 

                                        scriptProc.StartInfo.Arguments = "//B //Nologo vbscript.vbs";
                                        scriptProc.StartInfo.Arguments = string.Format("\"{0}\" \"{1}\" \"{2}\" \"{3}\" \"{4}\"",
                                            strMacroName, strMacroParm, strOutFolder, strOutputName + "." + strExportFormat, strLogFile);

                                        scriptProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden; //prevent console window from popping up

                                        try
                                        {
                                            scriptProc.Start();
                                            scriptProc.WaitForExit();  // <-- Optional if you want program running until your script exit

                                            int exitcode = scriptProc.ExitCode;
                                            scriptProc.Close();

                                            if (exitcode != 0)
                                                MessageBox.Show("Error executing macro " + strMacroName + ". Exit code is " + exitcode, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show("Error executing macro " + strMacroName + ". Error message is " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }

                                        myFileFuncs.WriteLine(strLogFile, "Processing complete.");
                                    }
                                }

                                if (blDebugMode) myFileFuncs.WriteLine(strLogFile, "Extracted " + strOutputName + " from " + strGISLayer + " for partner " + strPartner + ".");

                                // Delete the output table if not required
                                if ((strOutputFormat != strOutputType) && (strOutputType != ""))
                                {
                                    if (blDebugMode) myFileFuncs.WriteLine(strLogFile, "Deleting the output table ...");
                                    blResult = myArcMapFuncs.DeleteTable(strOutLayer, blDebugMode);
                                    if (!blResult)
                                    {
                                        myFileFuncs.WriteLine(strLogFile, "Error deleting the output table");
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
                                myFileFuncs.WriteLine(strLogFile, "There are no features selected in " + strGISLayer + ".");
                            }
                            //myFileFuncs.WriteLine(strLogFile, "Process complete for GIS layer " + strGISLayer + ", extracting layer " + strOutputName + " for partner " + strPartner + ".");
                            //}
                        }
                        else
                        {
                            myFileFuncs.WriteLine(strLogFile, "Skipping output = '" + strGISLayer + "' - not required.");
                        }

                        myFileFuncs.WriteLine(strLogFile, "Completed process " + intExtractCnt + " of " + intExtractTot + ".");
                        if (blDebugMode) myFileFuncs.WriteLine(strLogFile, "Process complete for GIS layer " + strGISLayer + " for partner " + strPartner + ".");
                        myFileFuncs.WriteLine(strLogFile, "");

                        intLayerIndex++;
                    }
                }

                // Zip up the results if required
                if (blCreateZip)
                {
                    lblStatus.Text = strPartner + ": Creating zip file.";
                    lblStatus.Refresh();
                    this.BringToFront();
                    string strSourcePath = strOutFolder;
                    string strOutPath = strOutFolder + ".zip";

                    // If a previous zip file exists, delete it.
                    string strPreviousZip = strOutFolder + @"\" + strOutputFolder + ".zip";
                    if (myFileFuncs.FileExists(strPreviousZip))
                        myFileFuncs.DeleteFile(strPreviousZip);

                    if (myFileFuncs.DirExists(strSourcePath))
                    {
                        
                        myFileFuncs.WriteLine(strLogFile, "Creating zip file " + strOutPath + "......");

                        ArchiveCreator AC = new WinBaseZipCreator(strOutPath);
                        ArchiveObj anObj = AC.GetArchive();
                        anObj.AddAllFiles(strSourcePath);

                        myFileFuncs.WriteLine(strLogFile, "Writing zip file...");
                        try
                        {
                            anObj.SaveArchive();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Could not write zip file for partner " + strPartner + ". Error message is " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            myFileFuncs.WriteLine(strLogFile, "Could not write zip file. Error message is " + ex.Message);
                            myFileFuncs.WriteLine(strLogFile, "Zip file not created for partner " + strPartner);
                            anObj = null;
                            AC = null;
                            if (myFileFuncs.FileExists(strOutFolder + ".zip"))
                                myFileFuncs.DeleteFile(strOutFolder + ".zip");
                        }
                        anObj = null;
                        AC = null;

                        // Move the zip file to its final location
                        myFileFuncs.WriteLine(strLogFile, "Moving zip file to final location: " + strOutPath + ".");
                        string strFinalOutPath = strOutFolder + @"\" + strOutputFolder + ".zip";

                        if (myFileFuncs.FileExists(strFinalOutPath))
                            myFileFuncs.DeleteFile(strFinalOutPath);

                        File.Move(strOutPath, strFinalOutPath);

                        myFileFuncs.WriteLine(strLogFile, "Zip file created successfully for partner " + strPartner + ".");
                        
                    }
                }
                
                // Log the completion of this partner.
                myFileFuncs.WriteLine(strLogFile, "Processing for partner complete.");
                myFileFuncs.WriteLine(strLogFile, "----------------------------------------------------------------------");
            }

            // Clear the selection on the partner layer
            myArcMapFuncs.ClearSelection(strPartnerTable);

            lblStatus.Text = "";
            lblStatus.Refresh();
            this.BringToFront();

            // Switch drawing back on.
            myArcMapFuncs.ToggleDrawing();
            myArcMapFuncs.ToggleTOC();
            this.BringToFront();
            this.Cursor = Cursors.Default;

            myFileFuncs.WriteLine(strLogFile, "");
            myFileFuncs.WriteLine(strLogFile, "----------------------------------------------------------------------");
            myFileFuncs.WriteLine(strLogFile, "Process completed!");
            myFileFuncs.WriteLine(strLogFile, "----------------------------------------------------------------------");

            DialogResult dlResult = MessageBox.Show("Process complete. Do you wish to close the form?", "Data Extractor", MessageBoxButtons.YesNo);
            if (dlResult == System.Windows.Forms.DialogResult.Yes)
                this.Close();
            else this.BringToFront();

            Process.Start("notepad.exe", strLogFile);
            return;
        }

        private void frmDataExtractor_Load(object sender, EventArgs e)
        {

        }

        private void lstActivePartners_DoubleClick(object sender, EventArgs e)
        {
            // Show the comment that is related to the selected item.
            if (lstActivePartners.SelectedItem != null)
            {
                // Get the partner name
                string strPartner = lstActivePartners.SelectedItem.ToString();

                // Get all the other partner details
                int a = liActivePartnerNames.IndexOf(strPartner);
                string strPartnerShort = liActivePartnerShorts[a];
                string strPartnerFormat = liActivePartnerFormats[a];
                string strPartnerExport = liActivePartnerExports[a];
                string strPartnerSQLTable = liActivePartnerSQLTables[a];
                string strPartnerNotes = liActivePartnerNotes[a];

                string strText = string.Format("{0} ({1})\r\n\r\nGIS Format : {2}\r\n\r\nExport Format : {3}\r\n\r\nSQL Table : {4}\r\n\r\nNotes : {5}",
                    strPartner, strPartnerShort, strPartnerFormat, strPartnerExport, strPartnerSQLTable, strPartnerNotes);
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

// The DataTools are a suite of ArcGIS Pro addins used to extract
// and manage biodiversity information from ArcGIS Pro and SQL Server
// based on pre-defined or user specified criteria.
//
// Copyright © 2024 Andy Foy Consulting.
//
// This file is part of DataTools suite of programs..
//
// DataTools are free software: you can redistribute it and/or modify
// them under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// DataTools are distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with with program.  If not, see <http://www.gnu.org/licenses/>.

using DataExtractor.UI;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Xml;
using MessageBox = ArcGIS.Desktop.Framework.Dialogs.MessageBox;

//This configuration file reader loads all of the variables to
// be used by the tool. Some are mandatory, the remainder optional.

namespace DataExtractor
{
    /// <summary>
    /// This class reads the config XML file and stores the results.
    /// </summary>
    internal class DataExtractorConfig
    {
        #region Fields

        private static string _toolName;

        // Initialise component to read XML
        private readonly XmlElement _xmlDataExtractor;

        #endregion Fields

        #region Constructor

        /// <summary>
        /// Load the XML profile and read the variables.
        /// </summary>
        /// <param name="xmlFile"></param>
        /// <param name="toolName"></param>
        /// <param name="msgErrors"></param>
        public DataExtractorConfig(string xmlFile, string toolName, bool msgErrors)
        {
            _toolName = toolName;

            // The user has specified the xmlFile and we've checked it exists.
            _xmlFound = true;
            _xmlLoaded = true;

            // Load the XML file into memory.
            XmlDocument xmlConfig = new();
            try
            {
                xmlConfig.Load(xmlFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading XML file. " + ex.Message, _toolName, MessageBoxButton.OK, MessageBoxImage.Error);
                _xmlLoaded = false;
                return;
            }

            // Get the InitialConfig node (the first node).
            XmlNode currNode = xmlConfig.DocumentElement.FirstChild;
            _xmlDataExtractor = (XmlElement)currNode;

            if (_xmlDataExtractor == null)
            {
                MessageBox.Show("Error loading XML file.", _toolName, MessageBoxButton.OK, MessageBoxImage.Error);
                _xmlLoaded = false;
                return;
            }

            // Get the mandatory variables.
            try
            {
                if (!GetMandatoryVariables())
                    return;
            }
            catch (Exception ex)
            {
                // Only report message if user was prompted for the XML
                // file (i.e. the user interface has already loaded).
                if (msgErrors)
                    MessageBox.Show("Error loading XML file. " + ex.Message, _toolName, MessageBoxButton.OK, MessageBoxImage.Error);
                _xmlLoaded = false;
                return;
            }
        }

        #endregion Constructor

        #region Get Mandatory Variables

        /// <summary>
        /// Get the mandatory variables from the XML file.
        /// </summary>
        /// <returns></returns>
        public bool GetMandatoryVariables()
        {
            string rawText;

            // The existing file location where log files will be saved with output messages.
            try
            {
                _logFilePath = _xmlDataExtractor["LogFilePath"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'LogFilePath' in the XML profile.");
            }

            // The location of the SDE file that specifies which SQL Server database to connect to.
            try
            {
                _sdeFile = _xmlDataExtractor["SDEFile"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'SDEFile' in the XML profile.");
            }

            // The schema used in the SQL Server database.
            try
            {
                _databaseSchema = _xmlDataExtractor["DatabaseSchema"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'DatabaseSchema' in the XML profile.");
            }

            // The stored procedure to execute spatial selection in SQL Server.
            try
            {
                _spatialStoredProcedure = _xmlDataExtractor["SpatialStoredProcedure"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'SpatialStoredProcedure' in the XML profile.");
            }

            // The stored procedure to execute non-spatial subset selection in SQL Server.
            try
            {
                _subsetStoredProcedure = _xmlDataExtractor["SubsetStoredProcedure"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'SubsetStoredProcedure' in the XML profile.");
            }

            // The stored procedure to clear the spatial selection in SQL Server.
            try
            {
                _clearSpatialStoredProcedure = _xmlDataExtractor["ClearSpatialStoredProcedure"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'ClearSpatialStoredProcedure' in the XML profile.");
            }

            // The stored procedure to clear the subset selection in SQL Server.
            try
            {
                _clearSubsetStoredProcedure = _xmlDataExtractor["ClearSubsetStoredProcedure"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'ClearSubsetStoredProcedure' in the XML profile.");
            }

            // The existing file location under which all partner sub-folders will be created.
            try
            {
                _defaultPath = _xmlDataExtractor["DefaultPath"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'DefaultPath' in the XML profile.");
            }

            // The name of the sub-folder which wil be created for all partner outputs.
            try
            {
                _partnerFolder = _xmlDataExtractor["PartnerFolder"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'PartnerFolder' in the XML profile.");
            }

            // The output filegeodatabase into which GDB files will be saved.
            try
            {
                _gdbName = _xmlDataExtractor["GDBName"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'GDBName' in the XML profile.");
            }

            // The name of the sub-folder which wil be created for all partner ArcGIS outputs.
            try
            {
                _arcGISFolder = _xmlDataExtractor["ArcGISFolder"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'ArcGISFolder' in the XML profile.");
            }

            // The name of the sub-folder which wil be created for all partner CSV outputs.
            try
            {
                _csvFolder = _xmlDataExtractor["CSVFolder"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'CSVFolder' in the XML profile.");
            }

            // The name of the sub-folder which wil be created for all partner TXT outputs.
            try
            {
                _txtFolder = _xmlDataExtractor["TXTFolder"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'TXTFolder' in the XML profile.");
            }

            // The schema used in the SQL Server database.
            try
            {
                _databaseSchema = _xmlDataExtractor["DatabaseSchema"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'DatabaseSchema' in the XML profile.");
            }

            // The Include wildcard for table names to list all the species tables
            // in SQL Server that can be selected by the user to extract from.
            try
            {
                _includeWildcard = _xmlDataExtractor["IncludeWildcard"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'IncludeWildcard' in the XML profile.");
            }

            // The Exclude wildcard for table names that should NOT be used
            // for species tables in SQL Server that can be selected by the
            // user to extract from.
            try
            {
                _excludeWildcard = _xmlDataExtractor["ExcludeWildcard"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'ExcludeWildcard' in the XML profile.");
            }

            // Whether the map processing should be paused during processing?
            try
            {
                _pauseMap = false;
                rawText = _xmlDataExtractor["PauseMap"].InnerText;
                if (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                    _pauseMap = true;
            }
            catch
            {
                // This is an optional node
                _pauseMap = false;
            }

            // The name of the partner GIS layer in SQL Server used to select the records.
            try
            {
                _partnerTable = _xmlDataExtractor["PartnerTable"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'PartnerTable' in the XML profile.");
            }

            // The name of the column in the partner GIS layer containing the partner
            // name passed to SQL Server by the tool to use as the partner's boundary
            // for selecting the records.
            try
            {
                _partnerColumn = _xmlDataExtractor["PartnerColumn"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'PartnerColumn' in the XML profile.");
            }

            // The name of the column in the partner GIS layer containing the
            // abbreviated name passed to SQL Server by the tool to use as the
            // sub-folder name for the destination of extracted.
            try
            {
                _shortColumn = _xmlDataExtractor["ShortColumn"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'ShortColumn' in the XML profile.");
            }

            // The name of the column in the partner GIS layer containing any
            // notes text relating to the partner.
            try
            {
                _notesColumn = _xmlDataExtractor["NotesColumn"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'NotesColumn' in the XML profile.");
            }

            // The name of the column in the partner GIS layer containing the
            // Y/N flag to indicate if the partner is currently active. Only
            // active partners will available for proccessing.
            try
            {
                _activeColumn = _xmlDataExtractor["ActiveColumn"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'ActiveColumn' in the XML profile.");
            }

            // The name of the column in the partner GIS layer containing the
            // GIS format required for the output records (SHP or GDB).
            try
            {
                _formatColumn = _xmlDataExtractor["FormatColumn"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'FormatColumn' in the XML profile.");
            }

            // The name of the column in the partner GIS layer indicating
            // whether an export should also be created as a CSV or TXT file.
            // Leave blank for no export.
            try
            {
                _exportColumn = _xmlDataExtractor["ExportColumn"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'ExportColumn' in the XML profile.");
            }

            // The name of the column in the partner GIS layer indicating
            // which SQL table should be used for each partner.
            try
            {
                _sqlTableColumn = _xmlDataExtractor["SQLTableColumn"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'SQLTableColumn' in the XML profile.");
            }

            // The name of the column in the partner GIS layer indicating
            // which SQL files should be created for each partner.
            try
            {
                _sqlFilesColumn = _xmlDataExtractor["SQLFilesColumn"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'SQLFilesColumn' in the XML profile.");
            }

            // The name of the column in the partner GIS layer indicating which
            // Map files should be created for each partner.
            try
            {
                _mapFilesColumn = _xmlDataExtractor["MapFilesColumn"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'MapFilesColumn' in the XML profile.");
            }

            // The name of the column in the partner GIS layer indicating which
            // survey tags, if any should be included in the export.
            try
            {
                _tagsColumn = _xmlDataExtractor["TagsColumn"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'TagsColumn' in the XML profile.");
            }

            // The name of the column in the partner GIS layer containing the spatial geometry.
            try
            {
                _spatialColumn = _xmlDataExtractor["SpatialColumn"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'SpatialColumn' in the XML profile.");
            }

            // The where clause to determine which partners to display.
            try
            {
                _partnerClause = _xmlDataExtractor["PartnerClause"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'PartnerClause' in the XML profile.");
            }

            // The options for the selection types. It is not recommended that these are changed.
            try
            {
                rawText = _xmlDataExtractor["SelectTypeOptions"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'SelectTypeOptions' in the XML profile.");
            }
            try
            {
                char[] chrSplitChars = [';'];
                _selectTypeOptions = [.. rawText.Split(chrSplitChars)];
            }
            catch
            {
                throw new("Error parsing 'BufferUnitOptions' string. Check for correct string formatting and placement of delimiters.");
            }

            // The default option (position in the list) to use for the selection types.
            try
            {
                rawText = _xmlDataExtractor["DefaultSelectType"].InnerText;
                bool blResult = Double.TryParse(rawText, out double i);
                if (blResult)
                    _defaultSelectType = (int)i;
                else
                {
                    throw new("The entry for 'DefaultSelectType' in the XML document is not a number.");
                }
            }
            catch
            {
                // This is an optional node
                _defaultSelectType = -1;
            }

            // The SQL criteria for excluding any unwanted records.
            try
            {
                _exclusionClause = _xmlDataExtractor["ExclusionClause"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'ExclusionClause' in the XML profile.");
            }

            // By default, should the exclusion clause be applied?
            try
            {
                _defaultApplyExclusionClause = false;
                rawText = _xmlDataExtractor["DefaultApplyExclusionClause"].InnerText;
                if (string.IsNullOrEmpty(rawText))
                    _defaultApplyExclusionClause = null;
                if (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                    _defaultApplyExclusionClause = true;
            }
            catch
            {
                // This is an optional node
                _defaultApplyExclusionClause = false;
            }

            // By default, should polygons be selected by centroids.
            try
            {
                _defaultUseCentroids = false;
                rawText = _xmlDataExtractor["DefaultUseCentroids"].InnerText;
                if (string.IsNullOrEmpty(rawText))
                    _defaultUseCentroids = null;
                if (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                    _defaultUseCentroids = true;
            }
            catch
            {
                // This is an optional node
                _defaultUseCentroids = false;
            }

            // By default, should the partner table be upload to the server.
            try
            {
                _defaultUploadToServer = false;
                rawText = _xmlDataExtractor["DefaultUploadToServer"].InnerText;
                if (string.IsNullOrEmpty(rawText))
                    _defaultUploadToServer = null;
                if (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                    _defaultUploadToServer = true;
            }
            catch
            {
                // This is an optional node
                _defaultUploadToServer = false;
            }

            // By default, should an existing log file be cleared?
            try
            {
                _defaultClearLogFile = false;
                rawText = _xmlDataExtractor["DefaultClearLogFile"].InnerText;
                if (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                    _defaultClearLogFile = true;
            }
            catch
            {
                // This is an optional node
                _defaultClearLogFile = false;
            }

            // By default, should the log file be opened after running?
            try
            {
                _defaultOpenLogFile = false;
                rawText = _xmlDataExtractor["DefaultOpenLogFile"].InnerText;
                if (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                    _defaultOpenLogFile = true;
            }
            catch
            {
                // This is an optional node
                _defaultOpenLogFile = false;
            }

            // All mandatory variables were loaded successfully.
            return true;
        }

        #endregion Get Mandatory Variables

        #region Get SQL Variables

        /// <summary>
        /// Get the SQL variables from the XML file.
        /// </summary>
        public bool GetSQLVariables()
        {
            // The SQL tables collection.
            XmlElement SQLTablesCollection;
            try
            {
                SQLTablesCollection = _xmlDataExtractor["SQLTables"];
            }
            catch
            {
                throw new("Could not locate the item 'SQLTables' in the XML profile");
            }

            // Reset the SQL layers list.
            _sqlLayers = [];

            // Now cycle through all of the maps.
            if (SQLTablesCollection != null)
            {
                foreach (XmlNode node in SQLTablesCollection)
                {
                    // Only process if not a comment
                    if (node.NodeType != XmlNodeType.Comment)
                    {
                        string nodeName = node.Name;
                        nodeName = nodeName.Replace("_", " "); // Replace any underscores with spaces for better display.

                        // Create a new layer for this node.
                        SQLLayer layer = new(nodeName);

                        try
                        {
                            string nodeGroup = nodeName.Substring(0, nodeName.IndexOf('-')).Trim();
                            string nodeTable = nodeName.Substring(nodeName.IndexOf('-') + 1).Trim();
                            layer.NodeGroup = nodeGroup;
                            layer.NodeTable = nodeTable;
                        }
                        catch
                        {
                            layer.NodeGroup = null;
                            layer.NodeTable = nodeName;
                        }

                        try
                        {
                            layer.OutputName = node["OutputName"].InnerText;
                        }
                        catch
                        {
                            throw new("Could not locate the item 'OutputName' for sql table " + nodeName + " in the XML file");
                        }

                        try
                        {
                            layer.Columns = node["Columns"].InnerText;
                        }
                        catch
                        {
                            throw new("Could not locate the item 'Columns' for sql table " + nodeName + " in the XML file");
                        }

                        try
                        {
                            layer.OutputType = node["OutputType"].InnerText;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.OutputType = null;
                        }

                        try
                        {
                            layer.WhereClause = node["WhereClause"].InnerText;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.WhereClause = null;
                        }

                        try
                        {
                            layer.OrderColumns = node["OrderColumns"].InnerText;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.OrderColumns = null;
                        }

                        try
                        {
                            layer.MacroName = node["MacroName"].InnerText;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.MacroName = null;
                        }

                        try
                        {
                            layer.MacroParms = node["MacroParms"].InnerText;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.MacroParms = null;
                        }

                        // Add the layer to the list of SQL tables.
                        SQLLayers.Add(layer);
                    }
                }
            }

            // All mandatory variables were loaded successfully.
            return true;
        }

        #endregion Get SQL Variables

        #region Get Map Variables

        /// <summary>
        /// Get the map variables from the XML file.
        /// </summary>
        public bool GetMapVariables()
        {
            string rawText;

            // The map tables collection.
            XmlElement MapTablesCollection;
            try
            {
                MapTablesCollection = _xmlDataExtractor["MapTables"];
            }
            catch
            {
                throw new("Could not locate the item 'MapTables' in the XML profile");
            }

            // Reset the map layers list.
            _mapLayers = [];

            // Now cycle through all of the maps.
            if (MapTablesCollection != null)
            {
                foreach (XmlNode node in MapTablesCollection)
                {
                    // Only process if not a comment
                    if (node.NodeType != XmlNodeType.Comment)
                    {
                        string nodeName = node.Name;

                        // Replace any underscores with spaces for better display.
                        nodeName = nodeName.Replace("_", " ");

                        // Create a new layer for this node.
                        MapLayer layer = new(nodeName);

                        try
                        {
                            string nodeGroup = nodeName.Substring(0, nodeName.IndexOf('-')).Trim();
                            string nodeLayer = nodeName.Substring(nodeName.IndexOf('-') + 1).Trim();
                            layer.NodeGroup = nodeGroup;
                            layer.NodeLayer = nodeLayer;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.NodeGroup = null;
                            layer.NodeLayer = nodeName;
                        }

                        try
                        {
                            layer.LayerName = node["LayerName"].InnerText;
                        }
                        catch
                        {
                            throw new("Could not locate the item 'LayerName' for map layer " + nodeName + " in the XML file");
                        }

                        try
                        {
                            layer.OutputName = node["OutputName"].InnerText;
                        }
                        catch
                        {
                            throw new("Could not locate the item 'OutputName' for map layer " + nodeName + " in the XML file");
                        }

                        try
                        {
                            layer.Columns = node["Columns"].InnerText;
                        }
                        catch
                        {
                            throw new("Could not locate the item 'Columns' for map layer " + nodeName + " in the XML file");
                        }

                        try
                        {
                            layer.OutputType = node["OutputType"].InnerText;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.OutputType = null;
                        }

                        try
                        {
                            layer.WhereClause = node["WhereClause"].InnerText;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.WhereClause = null;
                        }

                        try
                        {
                            layer.OrderColumns = node["OrderColumns"].InnerText;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.OrderColumns = null;
                        }

                        try
                        {
                            bool loadWarning = false;
                            rawText = node["LoadWarning"].InnerText;
                            if (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                                loadWarning = true;

                            layer.LoadWarning = loadWarning;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.LoadWarning = false;
                        }

                        try
                        {
                            layer.MacroName = node["MacroName"].InnerText;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.MacroName = null;
                        }

                        try
                        {
                            layer.MacroParms = node["MacroParms"].InnerText;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.MacroParms = null;
                        }

                        // Add the layer to the list of map layers.
                        MapLayers.Add(layer);
                    }
                }
            }

            // All mandatory variables were loaded successfully.
            return true;
        }

        #endregion Get Map Variables

        #region Members

        private readonly bool _xmlFound;

        /// <summary>
        /// Has the XML file been found.
        /// </summary>
        public bool XMLFound
        {
            get
            {
                return _xmlFound;
            }
        }

        private readonly bool _xmlLoaded;

        /// <summary>
        ///  Has the XML file been loaded.
        /// </summary>
        public bool XMLLoaded
        {
            get
            {
                return _xmlLoaded;
            }
        }

        #endregion Members

        #region Variables

        private string _logFilePath;

        public string LogFilePath
        {
            get { return _logFilePath; }
        }

        private string _sdeFile;

        public string SDEFile
        {
            get { return _sdeFile; }
        }

        private string _spatialStoredProcedure;

        public string SpatialStoredProcedure
        {
            get { return _spatialStoredProcedure; }
        }

        private string _subsetStoredProcedure;

        public string SubsetStoredProcedure
        {
            get { return _subsetStoredProcedure; }
        }

        private string _clearSpatialStoredProcedure;

        public string ClearSpatialStoredProcedure
        {
            get { return _clearSpatialStoredProcedure; }
        }

        private string _clearSubsetStoredProcedure;

        public string ClearSubsetStoredProcedure
        {
            get { return _clearSubsetStoredProcedure; }
        }

        private string _defaultPath;

        public string DefaultPath
        {
            get { return _defaultPath; }
        }

        private string _partnerFolder;

        public string PartnerFolder
        {
            get { return _partnerFolder; }
        }

        private string _gdbName;

        public string GDBName
        {
            get { return _gdbName; }
        }

        private string _arcGISFolder;

        public string ArcGISFolder
        {
            get { return _arcGISFolder; }
        }

        private string _csvFolder;

        public string CSVFolder
        {
            get { return _csvFolder; }
        }

        private string _txtFolder;

        public string TXTFolder
        {
            get { return _txtFolder; }
        }

        private string _databaseSchema;

        public string DatabaseSchema
        {
            get { return _databaseSchema; }
        }

        private string _includeWildcard;

        public string IncludeWildcard
        {
            get { return _includeWildcard; }
        }

        private string _excludeWildcard;

        public string ExcludeWildcard
        {
            get { return _excludeWildcard; }
        }

        private bool _pauseMap;

        public bool PauseMap
        {
            get { return _pauseMap; }
        }

        public List<string> AllPartnerColumns
        {
            get
            {
                List<string> allPartnerColumns = [];
                allPartnerColumns.Add(PartnerColumn);
                allPartnerColumns.Add(ShortColumn);
                allPartnerColumns.Add(NotesColumn);
                allPartnerColumns.Add(ActiveColumn);
                allPartnerColumns.Add(FormatColumn);
                allPartnerColumns.Add(ExportColumn);
                allPartnerColumns.Add(SQLTableColumn);
                allPartnerColumns.Add(SQLFilesColumn);
                allPartnerColumns.Add(MapFilesColumn);
                allPartnerColumns.Add(TagsColumn);

                return allPartnerColumns;
            }
        }

        private string _partnerTable;

        public string PartnerTable
        {
            get { return _partnerTable; }
        }

        private string _partnerColumn;

        public string PartnerColumn
        {
            get { return _partnerColumn; }
        }

        private string _shortColumn;

        public string ShortColumn
        {
            get { return _shortColumn; }
        }

        private string _notesColumn;

        public string NotesColumn
        {
            get { return _notesColumn; }
        }

        private string _activeColumn;

        public string ActiveColumn
        {
            get { return _activeColumn; }
        }

        private string _formatColumn;

        public string FormatColumn
        {
            get { return _formatColumn; }
        }

        private string _exportColumn;

        public string ExportColumn
        {
            get { return _exportColumn; }
        }

        private string _sqlTableColumn;

        public string SQLTableColumn
        {
            get { return _sqlTableColumn; }
        }

        private string _sqlFilesColumn;

        public string SQLFilesColumn
        {
            get { return _sqlFilesColumn; }
        }

        private string _mapFilesColumn;

        public string MapFilesColumn
        {
            get { return _mapFilesColumn; }
        }

        private string _tagsColumn;

        public string TagsColumn
        {
            get { return _tagsColumn; }
        }

        private string _spatialColumn;

        public string SpatialColumn
        {
            get { return _spatialColumn; }
        }

        private string _partnerClause;

        public string PartnerClause
        {
            get { return _partnerClause; }
        }

        private List<string> _selectTypeOptions = [];

        public List<string> SelectTypeOptions
        {
            get { return _selectTypeOptions; }
        }

        private int _defaultSelectType;

        public int DefaultSelectType
        {
            get { return _defaultSelectType; }
        }

        private string _exclusionClause;

        public string ExclusionClause
        {
            get { return _exclusionClause; }
        }

        private bool? _defaultApplyExclusionClause;

        public bool? DefaultApplyExclusionClause
        {
            get { return _defaultApplyExclusionClause; }
        }

        private bool? _defaultUseCentroids;

        public bool? DefaultUseCentroids
        {
            get { return _defaultUseCentroids; }
        }

        private bool? _defaultUploadToServer;

        public bool? DefaultUploadToServer
        {
            get { return _defaultUploadToServer; }
        }

        private bool _defaultClearLogFile;

        public bool DefaultClearLogFile
        {
            get { return _defaultClearLogFile; }
        }

        private bool _defaultOpenLogFile;

        public bool DefaultOpenLogFile
        {
            get { return _defaultOpenLogFile; }
        }

        #endregion Variables

        #region SQL Variables

        private List<SQLLayer> _sqlLayers = [];

        public List<SQLLayer> SQLLayers
        {
            get { return _sqlLayers; }
        }

        #endregion SQL Variables

        #region Map Variables

        private List<MapLayer> _mapLayers = [];

        public List<MapLayer> MapLayers
        {
            get { return _mapLayers; }
        }

        #endregion Map Variables
    }
}
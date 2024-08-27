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

using ArcGIS.Core.Data;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Internal.Framework.Controls;
using ArcGIS.Desktop.Mapping;
using DataTools;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using MessageBox = ArcGIS.Desktop.Framework.Dialogs.MessageBox;

namespace DataExtractor.UI
{
    internal class PaneHeader2ViewModel : PanelViewModelBase, INotifyPropertyChanged
    {
        #region Enums

        /// <summary>
        /// An enumeration of the different options for whether to
        /// add output layers to the map.
        /// </summary>
        public enum AddSelectedLayersOptions
        {
            No,
            WithoutLabels,
            WithLabels
        };

        /// <summary>
        /// An enumeration of the different options for whether to
        /// overwrite the label for each output layer.
        /// </summary>
        public enum OverwriteLabelOptions
        {
            No,
            ResetByLayer,
            ResetByGroup,
            DoNotReset
        };

        /// <summary>
        /// An enumeration of the different options for whether to
        /// overwrite the label for each output layer.
        /// </summary>
        public enum CombinedSitesTableOptions
        {
            None,
            Append,
            Overwrite
        };

        #endregion Enums

        #region Fields

        private readonly DockpaneMainViewModel _dockPane;

        private string _userID;
        private string _tempGDBName;

        private bool _searchErrors;

        private bool _updateTable;

        private string _repChar;

        private string _databasePath;
        private string _databaseTable;
        private string _databaseRefColumn;
        private string _databaseSiteColumn;
        private string _databaseOrgColumn;

        private bool _requireSiteName;
        private bool _requireOrganisation;

        private bool? _defaultKeepSelectedLayers;
        private int _defaultAddSelectedLayers;
        private int _defaultOverwriteLabels;
        private int _defaultCombinedSitesTable;

        private List<string> _bufferUnitOptionsDisplay;
        private List<string> _bufferUnitOptionsProcess;
        private List<string> _bufferUnitOptionsShort;

        private List<MapLayer> _mapLayers;

        private string _saveRootDir;
        private string _saveFolder;
        private string _gisFolder;
        private string _outputFolder;
        private string _layerFolder;

        private string _logFileName;
        private string _combinedSitesTableName;
        private string _combinedSitesColumnList;
        private string _combinedSitesTableFormat;
        private string _combinedSitesOutputFile;

        private string _bufferPrefix;
        private string _bufferLayerName;
        private string _bufferLayerPath;
        private string _bufferOutputFile;
        private string _bufferLayerFile;
        private string _bufferFields;
        private bool _keepBuffer;

        private string _searchLayerName;
        private string _searchOutputFile;
        private string _groupLayerName;

        private bool _keepSearchFeature;
        private string _searchSymbologyBase;

        private string _searchLayerBase;
        private List<string> _searchLayerExtensions;
        private string _searchLayerExtension;
        private string _inputLayerName;

        private string _searchColumn;
        private string _siteColumn;
        private string _orgColumn;
        private string _radiusColumn;

        private List<string> _mapGroupNames = [];
        private List<int> _mapGroupLabels = [];
        private int _maxLabel = 1;

        private string _logFile;

        private Geodatabase _tempGDB;
        private string _tempMasterLayerName;
        private string _tempMasterOutputFile;
        private string _tempFCLayerName;
        private string _tempFCOutputFile;
        private string _tempTableLayerName;
        private string _tempTableOutputFile;

        private string _tempFCPointsLayerName;
        private string _tempFCPointsOutputFile;
        private string _tempSearchPointsLayerName;
        private string _tempSearchPointsOutputFile;

        private const string _displayName = "DataExtractor";

        private readonly DataExtractorConfig _toolConfig;
        private MapFunctions _mapFunctions;

        #endregion Fields

        #region ViewModelBase Members

        public override string DisplayName
        {
            get { return _displayName; }
        }

        #endregion ViewModelBase Members

        #region Creator

        /// <summary>
        /// Set the global variables.
        /// </summary>
        /// <param name="xmlFilesList"></param>
        /// <param name="defaultXMLFile"></param>
        public PaneHeader2ViewModel(DockpaneMainViewModel dockPane, DataExtractorConfig toolConfig)
        {
            _dockPane = dockPane;

            // Return if no config object passed.
            if (toolConfig == null) return;

            // Set the config object.
            _toolConfig = toolConfig;

            InitializeComponent();
        }

        /// <summary>
        /// Initialise the search pane.
        /// </summary>
        private void InitializeComponent()
        {
            // Create a new map functions object.
            _mapFunctions = new();

            // Get the relevant config file settings.
            _logFilePath = _toolConfig.LogFilePath;

            _sdeFile = _toolConfig.SDEFile;

            _spatialStoredProcedure = _toolConfig.SpatialStoredProcedure;
            _subsetStoredProcedure = _toolConfig.SubsetStoredProcedure;
            _clearStoredProcedure = _toolConfig.ClearStoredProcedure;

            _defaultPath = _toolConfig.DefaultPath;
            _databaseSchema = _toolConfig.DatabaseSchema;

            _includeWildcard = _toolConfig.IncludeWildcard;
            _excludeWildcard = _toolConfig.ExcludeWildcard;

            _pauseMap = _toolConfig.PauseMap;

            _partnerTable = _toolConfig.PartnerTable;
            _partnerColumn = _toolConfig.PartnerColumn;
            _shortColumn = _toolConfig.ShortColumn;
            _notesColumn = _toolConfig.NotesColumn;
            _activeColumn = _toolConfig.ActiveColumn;
            _formatColumn = _toolConfig.FormatColumn;
            _exportColumn = _toolConfig.ExportColumn;
            _sqlFilesColumn = _toolConfig.SQLFilesColumn;
            _mapFilesColumn = _toolConfig.MapFilesColumn;
            _tagsColumn = _toolConfig.TagsColumn;
            _spatialColumn = _toolConfig.SpatialColumn;

            _selectTypeOptions = _toolConfig.SelectTypeOptions;
            _defaultSelectType; = _toolConfig.DefaultSelectType; ;

            _defaultConfidential = _toolConfig.DefaultConfidential;
            _defaultClearLogFile = _toolConfig.DefaultClearLogFile;
            _defaultOpenLogFile = _toolConfig.DefaultOpenLogFile;
            _defaultUseCentroid = _toolConfig.DefaultUseCentroid;

            // Get all of the SQL layer details.
            _sqlLayersList = _toolConfig.SQLLayers;

            // Get all of the map layer details.
            _mapLayersList = _toolConfig.MapLayers;
        }

        #endregion Creator

        #region Controls Enabled

        /// <summary>
        /// Is the list of layers enabled?
        /// </summary>
        public bool LayersListEnabled
        {
            get
            {
                return ((_dockPane.ProcessStatus == null)
                    && (_openLayersList != null));
            }
        }

        /// <summary>
        /// Is the list of buffer units options enabled?
        /// </summary>
        public bool BufferUnitsListEnabled
        {
            get
            {
                return ((_dockPane.ProcessStatus == null)
                    && (_bufferUnitsList != null));
            }
        }

        /// <summary>
        /// Is the option to keep the selected layers enabled?
        /// </summary>
        public bool KeepSelectedLayersEnabled
        {
            get
            {
                return ((_dockPane.ProcessStatus == null)
                    && (_defaultKeepSelectedLayers != null));
            }
        }

        /// <summary>
        /// Is the list of add to map options enabled?
        /// </summary>
        public bool AddToMapListEnabled
        {
            get
            {
                return ((_dockPane.ProcessStatus == null)
                    && (_addToMapList != null));
            }
        }

        /// <summary>
        /// Is the list of overwrite options enabled?
        /// </summary>
        public bool OverwriteLabelsListEnabled
        {
            get
            {
                return ((_dockPane.ProcessStatus == null)
                    && (_overwriteLabelsList != null));
            }
        }

        /// <summary>
        /// Is the list of combined sites options enabled?
        /// </summary>
        public bool CombinedSitesListEnabled
        {
            get
            {
                return ((_dockPane.ProcessStatus == null)
                    && (_combinedSitesList != null));
            }
        }

        /// <summary>
        /// Can the run button be pressed?
        /// </summary>
        public bool RunButtonEnabled
        {
            get
            {
                return ((_dockPane.ProcessStatus == null)
                    && (_openLayersList != null)
                    && (_openLayersList.Where(p => p.IsSelected).Any())
                    && (!string.IsNullOrEmpty(_searchRefText))
                    && (!_requireSiteName || !string.IsNullOrEmpty(_siteNameText))
                    && (!_requireOrganisation || !string.IsNullOrEmpty(_organisationText))
                    && (!string.IsNullOrEmpty(_bufferSizeText))
                    && (_selectedBufferUnitsIndex >= 0)
                    && (_defaultAddSelectedLayers <= 0 || _selectedAddToMap != null)
                    && (_defaultOverwriteLabels <= 0 || _selectedOverwriteLabels != null)
                    && (_defaultCombinedSitesTable <= 0 || _selectedCombinedSites != null));
            }
        }

        #endregion Controls Enabled

        #region Controls Visibility

        /// <summary>
        /// Is the site name text box visible.
        /// </summary>
        public Visibility SiteNameTextVisibility
        {
            get
            {
                if (!_requireSiteName)
                    return Visibility.Collapsed;
                else
                    return Visibility.Visible;
            }
        }

        /// <summary>
        /// Is the organisation text box visible.
        /// </summary>
        public Visibility OrganisationTextVisibility
        {
            get
            {
                if (!_requireOrganisation)
                    return Visibility.Collapsed;
                else
                    return Visibility.Visible;
            }
        }

        /// <summary>
        /// Is the option to keep output layers visible?
        /// </summary>
        public Visibility KeepLayersVisibility
        {
            get
            {
                if (_defaultKeepSelectedLayers == null)
                    return Visibility.Collapsed;
                else
                    return Visibility.Visible;
            }
        }

        /// <summary>
        /// Is the list of add to map options visible?
        /// </summary>
        public Visibility AddToMapListVisibility
        {
            get
            {
                if (_defaultKeepSelectedLayers == null || KeepSelectedLayers == false)
                    return Visibility.Collapsed;
                else
                {
                    if (_defaultAddSelectedLayers == -1)
                        return Visibility.Collapsed;
                    else
                        return Visibility.Visible;
                }
            }
        }

        /// <summary>
        /// Is the list of overwrite options visible?
        /// </summary>
        public Visibility OverwriteLabelsListVisibility
        {
            get
            {
                if (_defaultKeepSelectedLayers == null || KeepSelectedLayers == false)
                    return Visibility.Collapsed;
                else
                {
                    if (_defaultAddSelectedLayers == -1)
                        return Visibility.Collapsed;
                    else
                    {
                        try
                        {
                            int index = AddToMapList.FindIndex(a => a == SelectedAddToMap);
                            if (index > 0)
                                return Visibility.Visible;
                            else
                                return Visibility.Collapsed;
                        }
                        catch
                        {
                            return Visibility.Collapsed;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Is the list of combined sites options visible?
        /// </summary>
        public Visibility CombinedSitesListVisibility
        {
            get
            {
                if (_defaultCombinedSitesTable == -1)
                    return Visibility.Collapsed;
                else
                    return Visibility.Visible;
            }
        }

        #endregion Controls Visibility

        #region Message

        private string _message;

        /// <summary>
        /// The message to display on the form.
        /// </summary>
        public string Message
        {
            get
            {
                return _message;
            }
            set
            {
                _message = value;
                OnPropertyChanged(nameof(HasMessage));
                OnPropertyChanged(nameof(Message));
            }
        }

        private MessageType _messageLevel;

        /// <summary>
        /// The type of message; Error, Warning, Confirmation, Information
        /// </summary>
        public MessageType MessageLevel
        {
            get
            {
                return _messageLevel;
            }
            set
            {
                _messageLevel = value;
                OnPropertyChanged(nameof(MessageLevel));
            }
        }

        /// <summary>
        /// Is there a message to display?
        /// </summary>
        public Visibility HasMessage
        {
            get
            {
                if (_dockPane.ProcessStatus != null
                || string.IsNullOrEmpty(_message))
                    return Visibility.Collapsed;
                else
                    return Visibility.Visible;
            }
        }

        /// <summary>
        /// Show the message with the required icon (message type).
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="messageLevel"></param>
        public void ShowMessage(string msg, MessageType messageLevel)
        {
            MessageLevel = messageLevel;
            Message = msg;
        }

        /// <summary>
        /// Clear the form messages.
        /// </summary>
        public void ClearMessage()
        {
            Message = "";
        }

        #endregion Message

        #region Run Command

        /// <summary>
        /// Validates and executes the search.
        /// </summary>
        public async void RunSearch()
        {
            // Reset the cancel flag.
            _dockPane.SearchCancelled = false;

            // Validate the parameters.
            if (!ValidateParameters())
                return;

            // Create the search query.
            string searchRef = SearchRefText;
            string searchClause = _searchColumn + " = '" + searchRef + "'";

            // Find the features matching the search reference.
            if (!await FindSearchFeaturesAsync(searchClause))
            {
                ShowMessage("Search ref not found in any of the search layers.", MessageType.Warning);
                return;
            }

            // Clear any messages.
            ClearMessage();

            // Update the fields and buttons in the form.
            UpdateFormControls();
            _dockPane.RefreshPanel1Buttons();

            // Run the search.
            bool success = await RunSearchAsync();

            // Indicate that the search process has completed (successfully or not).
            string message;
            string image;
            if (success)
            {
                message = "Search '{0}' complete!";
                image = "Success";
            }
            else if (_searchErrors)
            {
                message = "Search '{0}' ended with errors!";
                image = "Error";
            }
            else if (_dockPane.SearchCancelled)
            {
                message = "Search '{0}' cancelled!";
                image = "Warning";
            }
            else
            {
                message = "Search '{0}' ended unexpectedly!";
                image = "Error";
            }

            StopSearch(searchRef, message, image);

            // Reset the cancel flag.
            _dockPane.SearchCancelled = false;

            // Update the fields and buttons in the form.
            UpdateFormControls();
            _dockPane.RefreshPanel1Buttons();
        }

        /// <summary>
        /// Validate the form parameters.
        /// </summary>
        /// <returns></returns>
        private bool ValidateParameters()
        {
            // Site ref is mandatory.
            if (string.IsNullOrEmpty(SearchRefText))
            {
                ShowMessage("Please enter a search reference.", MessageType.Warning);
                return false;
            }

            // Site name is not always required.
            if (_requireSiteName && string.IsNullOrEmpty(SiteNameText))
            {
                ShowMessage("Please enter a site name.", MessageType.Warning);
                return false;
            }

            // Organisation is not always required.
            if (_requireOrganisation && string.IsNullOrEmpty(OrganisationText))
            {
                ShowMessage("Please enter an organisation.", MessageType.Warning);
                return false;
            }

            // At least one layer must be selected,
            if (!OpenLayersList.Where(p => p.IsSelected).Any())
            {
                ShowMessage("Please select at least one layer to search.", MessageType.Warning);
                return false;
            }

            // A buffer size must be entered.
            if (string.IsNullOrEmpty(BufferSizeText))
            {
                ShowMessage("Please enter a buffer size.", MessageType.Warning);
                return false;
            }

            // The buffer suze must be numeric and positive.
            bool bufferNumeric = double.TryParse(BufferSizeText, out double bufferNumber);
            if (!bufferNumeric || bufferNumber < 0) // User either entered text or a negative number
            {
                ShowMessage("Please enter a positive number for the buffer size.", MessageType.Warning);
                return false;
            }

            // A buffer unit must be selected.
            if (SelectedBufferUnitsIndex < 0)
            {
                ShowMessage("Please select a buffer unit.", MessageType.Warning);
                return false;
            }

            // An add to map option must be selected (if visible).
            if (_defaultAddSelectedLayers != -1)
            {
                if (string.IsNullOrEmpty(SelectedAddToMap))
                {
                    ShowMessage("Please select whether layers should be added to the map.", MessageType.Warning);
                    return false;
                }
            }

            // An overwrite labels option must be selected (if visible).
            if (_defaultOverwriteLabels != -1 && !SelectedAddToMap.Equals("no", StringComparison.OrdinalIgnoreCase))
            {
                if (string.IsNullOrEmpty(SelectedOverwriteLabels))
                {
                    ShowMessage("Please select whether to overwrite labels for map layers.", MessageType.Warning);
                    return false;
                }
            }

            // A combined sites table option must be selected (if visible).
            if (_defaultCombinedSitesTable != -1)
            {
                if (string.IsNullOrEmpty(SelectedCombinedSites))
                {
                    ShowMessage("Please select whether the combined sites table should be created.", MessageType.Warning);
                    return false;
                }
            }

            ClearMessage();
            return true;
        }

        #endregion Run Command

        #region Properties

        /// <summary>
        /// Get the list of SQL layers.
        /// </summary>
        private ObservableCollection<SQLLayer> _sqlLayersList;

        /// <summary>
        /// Get the list of loaded GIS layers.
        /// </summary>
        public ObservableCollection<SQLLayer> SQLLayersList
        {
            get
            {
                return _sqlLayersList;
            }
        }

        /// <summary>
        /// Triggered  when the selection in the list of SQL layers changes.
        /// </summary>
        public int SQLLayersList_SelectedIndex
        {
            set
            {
                // Check if the run button is now enabled/disabled.
                _dockPane.CheckRunButton();
            }
        }

        private List<SQLLayer> _selectedSQLLayers;

        /// <summary>
        /// Get/Set the selected SQL layers.
        /// </summary>
        public List<SQLLayer> SelectedSQLLayers
        {
            get
            {
                return _selectedSQLLayers;
            }
            set
            {
                _selectedSQLLayers = value;
            }
        }

        /// <summary>
        /// Get the list of GIS Map layers.
        /// </summary>
        private ObservableCollection<MapLayer> _mapLayersList;

        /// <summary>
        /// Get the list of loaded GIS Map layers.
        /// </summary>
        public ObservableCollection<MapLayer> MapLayersList
        {
            get
            {
                return _mapLayersList;
            }
        }

        /// <summary>
        /// Triggered  when the selection in the list of GIS Map layers changes.
        /// </summary>
        public int MapLayersList_SelectedIndex
        {
            set
            {
                // Check if the run button is now enabled/disabled.
                _dockPane.CheckRunButton();
            }
        }

        private List<MapLayer> _selectedMapLayers;

        /// <summary>
        /// Get/Set the selected GIS Map layers.
        /// </summary>
        public List<MapLayer> SelectedMapLayers
        {
            get
            {
                return _selectedMapLayers;
            }
            set
            {
                _selectedMapLayers = value;
            }
        }



        private string _bufferSizeText;

        /// <summary>
        /// Get/Set the buffer size text.
        /// </summary>
        public string BufferSizeText
        {
            get
            {
                return _bufferSizeText;
            }
            set
            {
                _bufferSizeText = value;

                // Check if the run button is now enabled/disabled.
                _dockPane.CheckRunButton();
            }
        }

        private List<string> _bufferUnitsList;

        /// <summary>
        /// Get/Set the buffer size units.
        /// </summary>
        public List<string> BufferUnitsList
        {
            get
            {
                return _bufferUnitsList;
            }
            set
            {
                _bufferUnitsList = value;

                // Update the fields and buttons in the form.
                OnPropertyChanged(nameof(BufferUnitsListEnabled));
            }
        }

        private int _selectedBufferUnitsIndex;

        /// <summary>
        /// Get/Set the selected buffer size units index.
        /// </summary>
        public int SelectedBufferUnitsIndex
        {
            get
            {
                return _selectedBufferUnitsIndex;
            }
            set
            {
                _selectedBufferUnitsIndex = value;

                // Check if the run button is now enabled/disabled.
                _dockPane.CheckRunButton();
            }
        }

        private bool _keepSelectedLayers;

        /// <summary>
        /// Get/Set the option for whether to keep the selected output
        /// layers.
        /// </summary>
        public bool KeepSelectedLayers
        {
            get
            {
                return _keepSelectedLayers;
            }
            set
            {
                _keepSelectedLayers = value;

                // Update the fields and buttons in the form.
                OnPropertyChanged(nameof(KeepSelectedLayersEnabled));
                OnPropertyChanged(nameof(AddToMapListVisibility));
                OnPropertyChanged(nameof(OverwriteLabelsListVisibility));
            }
        }

        private List<string> _addToMapList;

        /// <summary>
        /// Get/Set the options for whether the output layers will be added
        /// to the map.
        /// </summary>
        public List<string> AddToMapList
        {
            get
            {
                return _addToMapList;
            }
            set
            {
                _addToMapList = value;

                // Update the fields and buttons in the form.
                OnPropertyChanged(nameof(AddToMapListEnabled));
            }
        }

        private string _selectedAddToMap;

        /// <summary>
        /// Get/Set whether the output layers will be added to the map.
        /// </summary>
        public string SelectedAddToMap
        {
            get
            {
                return _selectedAddToMap;
            }
            set
            {
                _selectedAddToMap = value;

                // Update the fields and buttons in the form.
                OnPropertyChanged(nameof(OverwriteLabelsListVisibility));

                // Check if the run button is now enabled/disabled.
                _dockPane.CheckRunButton();
            }
        }

        private List<string> _overwriteLabelsList;

        /// <summary>
        /// Get/Set the search site name.
        /// </summary>
        public List<string> OverwriteLabelsList
        {
            get
            {
                return _overwriteLabelsList;
            }
            set
            {
                _overwriteLabelsList = value;

                // Update the fields and buttons in the form.
                OnPropertyChanged(nameof(OverwriteLabelsListEnabled));
            }
        }

        private string _selectedOverwriteLabels;

        /// <summary>
        /// Get/Set the search site name.
        /// </summary>
        public string SelectedOverwriteLabels
        {
            get
            {
                return _selectedOverwriteLabels;
            }
            set
            {
                _selectedOverwriteLabels = value;

                // Check if the run button is now enabled/disabled.
                _dockPane.CheckRunButton();
            }
        }

        private List<string> _combinedSitesList;

        /// <summary>
        /// Get/Set the search site name.
        /// </summary>
        public List<string> CombinedSitesList
        {
            get
            {
                return _combinedSitesList;
            }
            set
            {
                _combinedSitesList = value;

                // Update the fields and buttons in the form.
                OnPropertyChanged(nameof(CombinedSitesListEnabled));
            }
        }

        private string _selectedCombinedSites;

        /// <summary>
        /// Get/Set the search site name.
        /// </summary>
        public string SelectedCombinedSites
        {
            get
            {
                return _selectedCombinedSites;
            }
            set
            {
                _selectedCombinedSites = value;

                // Check if the run button is now enabled/disabled.
                _dockPane.CheckRunButton();
            }
        }

        private bool _clearLogFile;

        /// <summary>
        /// Is the log file to be cleared before running the search?
        /// </summary>
        public bool ClearLogFile
        {
            get
            {
                return _clearLogFile;
            }
            set
            {
                _clearLogFile = value;
            }
        }

        private bool _openLogFile;

        /// <summary>
        /// Is the log file to be opened after running the search?
        /// </summary>
        public bool OpenLogFile
        {
            get
            {
                return _openLogFile;
            }
            set
            {
                _openLogFile = value;
            }
        }

        private bool _pauseMap;

        /// <summary>
        /// Whether the map processing should be paused during processing?
        /// </summary>
        public bool PauseMap
        {
            get
            {
                return _pauseMap;
            }
            set
            {
                _pauseMap = value;
            }
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Update the fields and buttons in the form.
        /// </summary>
        private void UpdateFormControls()
        {
            UpdateFormFields();

            // Check if the run button is now enabled/disabled.
            _dockPane.CheckRunButton();
        }

        /// <summary>
        /// Update the fields in the form.
        /// </summary>
        private void UpdateFormFields()
        {
            OnPropertyChanged(nameof(SearchRefText));
            OnPropertyChanged(nameof(SiteNameText));
            OnPropertyChanged(nameof(SiteNameTextVisibility));
            OnPropertyChanged(nameof(OrganisationText));
            OnPropertyChanged(nameof(OrganisationTextVisibility));
            OnPropertyChanged(nameof(OpenLayersList));
            OnPropertyChanged(nameof(LayersListEnabled));
            OnPropertyChanged(nameof(BufferSizeText));
            OnPropertyChanged(nameof(BufferUnitsList));
            OnPropertyChanged(nameof(BufferUnitsListEnabled));
            OnPropertyChanged(nameof(SelectedBufferUnitsIndex));
            OnPropertyChanged(nameof(AddToMapList));
            OnPropertyChanged(nameof(AddToMapListEnabled));
            OnPropertyChanged(nameof(SelectedAddToMap));
            OnPropertyChanged(nameof(AddToMapListVisibility));
            OnPropertyChanged(nameof(OverwriteLabelsList));
            OnPropertyChanged(nameof(OverwriteLabelsListEnabled));
            OnPropertyChanged(nameof(SelectedOverwriteLabels));
            OnPropertyChanged(nameof(OverwriteLabelsListVisibility));
            OnPropertyChanged(nameof(CombinedSitesList));
            OnPropertyChanged(nameof(SelectedCombinedSites));
            OnPropertyChanged(nameof(CombinedSitesListVisibility));
        }

        /// <summary>
        /// Lookup the search reference in the database.
        /// </summary>
        /// <param name="searchRefText"></param>
        /// <param name="siteName"></param>
        /// <param name="organisation"></param>
        /// <returns></returns>
        public bool LookupSearchRef(string searchRefText, ref string siteName, ref string organisation)
        {
            // Use connection string for .accdb or .mdb as appropriate.
            string connectionString;
            if (FileFunctions.GetExtension(_databasePath).Equals(".accdb", StringComparison.OrdinalIgnoreCase))
                connectionString = "Provider='Microsoft.ACE.OLEDB.12.0';data source='" + _databasePath + "'";
            else
                connectionString = "Provider='Microsoft.Jet.OLEDB.4.0';data source='" + _databasePath + "'";

            // Build the list of columns to retrieve.
            string mapColumns = _databaseSiteColumn;
            if (!string.IsNullOrEmpty(_databaseOrgColumn))
                mapColumns += "," + _databaseOrgColumn;

            // Build the search query.
            string searchQuery = "SELECT " + mapColumns + " FROM " + _databaseTable + " WHERE LCASE(" + _databaseRefColumn + ") = " + '"' + searchRefText + '"';

            siteName = null;
            organisation = null;

            // Connect to the database.
            OleDbConnection oleDbConnection;
            try
            {
                oleDbConnection = new(connectionString);
            }
            catch (Exception)
            {
                //MessageBox.Show("Error: Failed to create a database connection");
                return false;
            }

            DataSet dataSet = new();
            try
            {
                // Create an adapter.
                OleDbCommand oleDbCommand = new(searchQuery, oleDbConnection);
                OleDbDataAdapter dataAdapter = new(oleDbCommand);

                // Open the database.
                oleDbConnection.Open();

                // File the adapter for the table.
                dataAdapter.Fill(dataSet, _databaseTable);

                dataAdapter.Dispose();
                oleDbCommand.Dispose();
            }
            catch (Exception)
            {
                //MessageBox.Show("Error: Failed to retrieve the required data from the database");
                return false;
            }
            finally
            {
                oleDbConnection.Dispose();
                oleDbConnection.Close();
            }

            try
            {
                DataRowCollection dataRowCollection = dataSet.Tables[_databaseTable].Rows;
                foreach (DataRow aRow in dataRowCollection) // Really there should only be one. We can check for this.
                {
                    // Get the site name and organisation names.
                    siteName = aRow[0].ToString();
                    if (!string.IsNullOrEmpty(_databaseOrgColumn))
                        organisation = aRow[1].ToString();
                }
            }
            catch
            {
                //MessageBox.Show("Error: Failed to find the required table in the table");
                return false;
            }
            finally
            {
                dataSet.Dispose();
            }

            return true;
        }

        /// <summary>
        /// Set all of the form fields to their default values.
        /// </summary>
        public async Task ResetForm(bool reset)
        {
            // Clear the selections first (to avoid selections being retained).
            if (_openLayersList != null)
            {
                foreach (MapLayer layer in _openLayersList)
                {
                    layer.IsSelected = false;
                }
            }

            // Search ref and site name.
            SearchRefText = null;
            SiteNameText = null;
            OrganisationText = null;

            // Buffer size and units.
            BufferSizeText = _toolConfig.DefaultBufferSize.ToString();
            BufferUnitsList = _toolConfig.BufferUnitOptionsDisplay;
            if (_toolConfig.DefaultBufferUnit > 0)
                SelectedBufferUnitsIndex = _toolConfig.DefaultBufferUnit - 1;

            // Add layers to map.
            AddToMapList = _toolConfig.AddSelectedLayersOptions;
            if (_toolConfig.DefaultAddSelectedLayers > 0)
                SelectedAddToMap = _toolConfig.AddSelectedLayersOptions[_toolConfig.DefaultAddSelectedLayers - 1];

            // Overwrite map layers.
            OverwriteLabelsList = _toolConfig.OverwriteLabelOptions;
            if (_toolConfig.DefaultOverwriteLabels > 0)
                SelectedOverwriteLabels = _toolConfig.OverwriteLabelOptions[_toolConfig.DefaultOverwriteLabels - 1];

            // Combined sites table.
            CombinedSitesList = _toolConfig.CombinedSitesTableOptions;
            if (_toolConfig.DefaultCombinedSitesTable > 0)
                SelectedCombinedSites = _toolConfig.CombinedSitesTableOptions[_toolConfig.DefaultCombinedSitesTable - 1];

            // Log file.
            ClearLogFile = _toolConfig.DefaultClearLogFile;
            OpenLogFile = _toolConfig.DefaultOpenLogFile;

            // Pause map.
            PauseMap = _toolConfig.PauseMap;

            // Reload the form layers (don't wait for the response).
            await LoadLayersAsync(reset, true);
        }

        /// <summary>
        /// Load the list of open GIS layers.
        /// </summary>
        /// <param name="selectedTable"></param>
        /// <returns></returns>
        public async Task LoadLayersAsync(bool reset, bool message)
        {
            // If not already processing.
            if (_dockPane.ProcessStatus == null)
            {
                _dockPane.LayersListLoading = true;
                if (reset)
                    _dockPane.ProgressUpdate("Resetting form...");
                else
                    _dockPane.ProgressUpdate("Loading form...");

                // Clear any messages.
                ClearMessage();

                // Update the fields and buttons in the form.
                UpdateFormControls();

                List<string> closedLayers = []; // The closed layers by name.

                await Task.Run(() =>
                {
                    if (_mapFunctions == null || _mapFunctions.MapName == null || MapView.Active.Map.Name != _mapFunctions.MapName)
                    {
                        // Create a new map functions object.
                        _mapFunctions = new();
                    }

                    // Check if there is an active map.
                    bool mapOpen = _mapFunctions.MapName != null;

                    // Reset the list of open layers.
                    ObservableCollection<MapLayer> openLayersList = [];

                    if (mapOpen)
                    {
                        List<MapLayer> allLayers = _mapLayers;

                        // Loop through all of the layers to check if they are open
                        // in the active map.
                        foreach (MapLayer layer in allLayers)
                        {
                            if (_mapFunctions.FindLayer(layer.LayerName) != null)
                            {
                                // Preselect layer if required.
                                layer.IsSelected = layer.PreselectLayer;

                                // Add the open layers to the list.
                                openLayersList.Add(layer);
                            }
                            else
                            {
                                // Only add if the user wants to be warned of this one.
                                if (layer.LoadWarning)
                                    closedLayers.Add(layer.LayerName);
                            }
                        }
                    }

                    // Reset the list of open layers.
                    _openLayersList = openLayersList;
                });

                // Hide progress update.
                _dockPane.ProgressUpdate(null, -1, -1);

                // Show a message if there are no open layers.
                if (!_openLayersList.Any())
                    ShowMessage("No search layers in active map.", MessageType.Warning);

                _dockPane.LayersListLoading = false;

                // Update the fields and buttons in the form.
                UpdateFormControls();

                // Warn the user of closed layers.
                int closedLayerCount = closedLayers.Count;
                if (closedLayerCount > 0)
                {
                    string closedLayerWarning = "";
                    if (closedLayerCount == 1)
                    {
                        closedLayerWarning = "Layer '" + closedLayers[0] + "' is not loaded.";
                    }
                    else
                    {
                        closedLayerWarning = string.Format("{0} layers are not loaded.", closedLayerCount.ToString());
                    }

                    ShowMessage(closedLayerWarning, MessageType.Warning);

                    if (message)
                        MessageBox.Show(closedLayerWarning, "Data Searches", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }

        /// <summary>
        /// Clear the list of open GIS layers.
        /// </summary>
        /// <param name="selectedTable"></param>
        /// <returns></returns>
        public void ClearLayers()
        {
            // Clear the list of open layers.
            _openLayersList = [];

            // Update the fields and buttons in the form.
            UpdateFormControls();
        }

        /// <summary>
        /// Validate and run the search.
        /// </summary>
        private async Task<bool> RunSearchAsync()
        {
            if (_mapFunctions == null || _mapFunctions.MapName == null)
            {
                // Create a new map functions object.
                _mapFunctions = new();
            }

            // Reset search errors flag.
            _searchErrors = false;

            // Save the parameters.
            string searchRef = SearchRefText;
            string siteName = SiteNameText;
            string organisation = OrganisationText;
            string bufferSize = BufferSizeText;
            int bufferUnitIndex = SelectedBufferUnitsIndex;

            // Selected layers.
            _selectedLayers = OpenLayersList.Where(p => p.IsSelected).ToList();

            // What is the selected buffer unit?
            string bufferUnitText = _bufferUnitOptionsDisplay[bufferUnitIndex]; // Unit to be used in reporting.
            string bufferUnitProcess = _bufferUnitOptionsProcess[bufferUnitIndex]; // Unit to be used in process (because of American spelling).
            string bufferUnitShort = _bufferUnitOptionsShort[bufferUnitIndex]; // Unit to be used in file naming (abbreviation).

            // What is the area measurement unit?
            string areaMeasureUnit = _toolConfig.AreaMeasurementUnit;

            // Will the selected layers be kept?
            bool keepSelectedLayers = KeepSelectedLayers;

            // Will the selected layers be added to the map with labels?
            AddSelectedLayersOptions addSelectedLayersOption = AddSelectedLayersOptions.No;
            if (_defaultAddSelectedLayers > 0)
            {
                if (SelectedAddToMap.Equals("no", StringComparison.OrdinalIgnoreCase))
                    addSelectedLayersOption = AddSelectedLayersOptions.No;
                else if (SelectedAddToMap.Contains("with labels", StringComparison.OrdinalIgnoreCase))
                    addSelectedLayersOption = AddSelectedLayersOptions.WithLabels;
                else if (SelectedAddToMap.Contains("without labels", StringComparison.OrdinalIgnoreCase))
                    addSelectedLayersOption = AddSelectedLayersOptions.WithoutLabels;
            }

            // Will the labels on map layers be overwritten?
            OverwriteLabelOptions overwriteLabelOption = OverwriteLabelOptions.No;
            if (_defaultOverwriteLabels > 0)
            {
                if (SelectedOverwriteLabels.Equals("no", StringComparison.OrdinalIgnoreCase))
                    overwriteLabelOption = OverwriteLabelOptions.No;
                else if (SelectedOverwriteLabels.Contains("reset each layer", StringComparison.OrdinalIgnoreCase))
                    overwriteLabelOption = OverwriteLabelOptions.ResetByLayer;
                else if (SelectedOverwriteLabels.Contains("reset each group", StringComparison.OrdinalIgnoreCase))
                    overwriteLabelOption = OverwriteLabelOptions.ResetByGroup;
                else if (SelectedOverwriteLabels.Contains("do not reset", StringComparison.OrdinalIgnoreCase))
                    overwriteLabelOption = OverwriteLabelOptions.DoNotReset;
            }

            // Will the combined sites table be created and overwritten?
            CombinedSitesTableOptions combinedSitesTableOption = CombinedSitesTableOptions.None;
            if (_defaultCombinedSitesTable > 0)
            {
                if (SelectedCombinedSites.Equals("none", StringComparison.OrdinalIgnoreCase))
                    combinedSitesTableOption = CombinedSitesTableOptions.None;
                else if (SelectedCombinedSites.Contains("append", StringComparison.OrdinalIgnoreCase))
                    combinedSitesTableOption = CombinedSitesTableOptions.Append;
                else if (SelectedCombinedSites.Contains("overwrite", StringComparison.OrdinalIgnoreCase))
                    combinedSitesTableOption = CombinedSitesTableOptions.Overwrite;
            }

            // Fix any illegal characters in the site name string.
            siteName = StringFunctions.StripIllegals(siteName, _repChar);

            // Create the ref string from the search reference.
            string reference = searchRef.Replace("/", _repChar);

            // Create the shortref from the search reference by
            // getting rid of any characters.
            string shortRef = StringFunctions.KeepNumbersAndSpaces(reference, _repChar);

            // Find the subref part of this reference.
            string subref = StringFunctions.GetSubref(shortRef, _repChar);

            // Create the radius from the buffer size and units
            string radius = bufferSize + bufferUnitShort;

            // Replace any standard strings in the variables.
            _saveFolder = StringFunctions.ReplaceSearchStrings(_toolConfig.SaveFolder, reference, siteName, shortRef, subref, radius);
            _gisFolder = StringFunctions.ReplaceSearchStrings(_toolConfig.GISFolder, reference, siteName, shortRef, subref, radius);
            _logFileName = StringFunctions.ReplaceSearchStrings(_toolConfig.LogFileName, reference, siteName, shortRef, subref, radius);
            _combinedSitesTableName = StringFunctions.ReplaceSearchStrings(_toolConfig.CombinedSitesTableName, reference, siteName, shortRef, subref, radius);
            _bufferPrefix = StringFunctions.ReplaceSearchStrings(_toolConfig.BufferPrefix, reference, siteName, shortRef, subref, radius);
            _searchLayerName = StringFunctions.ReplaceSearchStrings(_toolConfig.SearchOutputName, reference, siteName, shortRef, subref, radius);
            _groupLayerName = StringFunctions.ReplaceSearchStrings(_toolConfig.GroupLayerName, reference, siteName, shortRef, subref, radius);

            // Remove any illegal characters from the names.
            _saveFolder = StringFunctions.StripIllegals(_saveFolder, _repChar);
            _gisFolder = StringFunctions.StripIllegals(_gisFolder, _repChar);
            _logFileName = StringFunctions.StripIllegals(_logFileName, _repChar, true);
            _combinedSitesTableName = StringFunctions.StripIllegals(_combinedSitesTableName, _repChar);
            _bufferPrefix = StringFunctions.StripIllegals(_bufferPrefix, _repChar);
            _searchLayerName = StringFunctions.StripIllegals(_searchLayerName, _repChar);
            _groupLayerName = StringFunctions.StripIllegals(_groupLayerName, _repChar);

            // Trim any trailing spaces (directory functions don't deal with them well).
            _saveFolder = _saveFolder.Trim();

            // Create output folders if required.
            _outputFolder = CreateOutputFolders(_saveRootDir, _saveFolder, _gisFolder);
            if (_outputFolder == null)
            {
                MessageBox.Show("Cannot create output folder");
                return false;
            }

            // Create log file (if necessary).
            _logFile = _outputFolder + @"\" + _logFileName;
            if (FileFunctions.FileExists(_logFile) && ClearLogFile)
            {
                try
                {
                    File.Delete(_logFile);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Cannot clear log file " + _logFile + ". Please make sure this file is not open in another window. " +
                        "System error: " + ex.Message);
                    return false;
                }
            }

            // Replace any illegal characters in the user name string.
            _userID = StringFunctions.StripIllegals(Environment.UserName, "_", false);

            // User ID should be something at least.
            if (string.IsNullOrEmpty(_userID))
            {
                _userID = "Temp";
                FileFunctions.WriteLine(_logFile, "User ID not found. User ID used will be 'Temp'");
            }

            // Count the number of layers to process and add 2
            // to account for the start and finish steps.
            int stepsMax = SelectedLayers.Count + 2;
            int stepNum = 0;

            // Stop if the user cancelled the process.
            if (_dockPane.SearchCancelled)
                return false;

            // Indicate the search has started.
            _dockPane.SearchCancelled = false;
            _dockPane.SearchRunning = true;

            // Write the first line to the log file.
            FileFunctions.WriteLine(_logFile, "-----------------------------------------------------------------------");
            FileFunctions.WriteLine(_logFile, "Processing search '" + searchRef + "'");
            FileFunctions.WriteLine(_logFile, "-----------------------------------------------------------------------");

            FileFunctions.WriteLine(_logFile, "Parameters are as follows:");
            FileFunctions.WriteLine(_logFile, "Buffer distance: " + radius);
            FileFunctions.WriteLine(_logFile, "Output location: " + _saveRootDir + @"\" + _saveFolder);
            FileFunctions.WriteLine(_logFile, "Layers to process: " + SelectedLayers.Count.ToString());
            FileFunctions.WriteLine(_logFile, "Area measurement unit: " + areaMeasureUnit);

            // Create the search query.
            string searchClause = _searchColumn + " = '" + searchRef + "'";

            _dockPane.ProgressUpdate("Selecting feature(s)...", stepNum, stepsMax);
            stepNum += 1;

            // Count the features matching the search reference.
            if (await CountSearchFeaturesAsync(searchClause) == 0)
            {
                _searchErrors = true;
                return false;
            }

            // Prepare the temporary geodatabase
            if (!await PrepareTemporaryGDBAsync())
            {
                _searchErrors = true;
                return false;
            }

            // Pause the map redrawing.
            if (PauseMap)
                _mapFunctions.PauseDrawing(true);

            // Select the feature matching the search reference in the map.
            if (!await _mapFunctions.SelectLayerByAttributesAsync(_inputLayerName, searchClause, SelectionCombinationMethod.New))
            {
                _searchErrors = true;
                return false;
            }

            // Update the table if required.
            if (_updateTable && (!string.IsNullOrEmpty(_siteColumn) || !string.IsNullOrEmpty(_orgColumn) || !string.IsNullOrEmpty(_radiusColumn)))
            {
                FileFunctions.WriteLine(_logFile, "Updating attributes in search layer ...");

                if (!await _mapFunctions.UpdateFeaturesAsync(_inputLayerName, _siteColumn, siteName, _orgColumn, organisation, _radiusColumn, radius))
                {
                    _searchErrors = true;
                    return false;
                }
            }

            // The output file for the search features is a shapefile in the root save directory.
            _searchOutputFile = _outputFolder + "\\" + _searchLayerName + ".shp";

            // Remove the search feature layer from the map
            // in case there is one already there from a different folder.
            await _mapFunctions.RemoveLayerAsync(_searchLayerName);

            // Save the selected feature(s).
            if (!await SaveSearchFeaturesAsync())
            {
                _searchErrors = true;
                return false;
            }

            // Stop if the user cancelled the process.
            if (_dockPane.SearchCancelled)
                return false;

            _dockPane.ProgressUpdate("Buffering feature(s)...", stepNum, stepsMax);
            stepNum += 1;

            // Set the buffer layer name by appending the radius.
            _bufferLayerName = _bufferPrefix + "_" + radius;
            if (_bufferLayerName.Contains('.'))
                _bufferLayerName = _bufferLayerName.Replace('.', '_');

            // The output file for the buffer is a shapefile in the root save directory.
            _bufferOutputFile = _outputFolder + "\\" + _bufferLayerName + ".shp";

            // Remove the buffer layer from the map
            // in case there is one already there from a different folder.
            await _mapFunctions.RemoveLayerAsync(_bufferLayerName);

            // Buffer search feature(s).
            if (!await BufferSearchFeaturesAsync(bufferSize, bufferUnitProcess, bufferUnitShort))
            {
                _searchErrors = true;
                return false;
            }

            // Zoom to the buffer layer extent or a fixed scale if no buffer.
            if (!_pauseMap)
            {
                if (bufferSize == "0")
                    await _mapFunctions.ZoomToLayerAsync(_searchLayerName, 1, 10000);
                else
                    await _mapFunctions.ZoomToLayerAsync(_bufferLayerName, 1.05);
            }

            // Get the full layer path (in case it's nested in one or more groups).
            _bufferLayerPath = _mapFunctions.GetLayerPath(_bufferLayerName);

            // Start the combined sites table before we do any analysis.
            _combinedSitesOutputFile = _outputFolder + @"\" + _combinedSitesTableName + "." + _combinedSitesTableFormat;
            if (!CreateCombinedSitesTable(_combinedSitesOutputFile, combinedSitesTableOption))
            {
                _searchErrors = true;
                return false;
            }

            // Get any groups and initialise required layers.
            if (overwriteLabelOption == OverwriteLabelOptions.ResetByLayer ||
                overwriteLabelOption == OverwriteLabelOptions.ResetByGroup)
            {
                _mapGroupNames = _selectedLayers.Select(l => l.NodeGroup).Distinct().ToList();
                _mapGroupLabels = [];
                foreach (string groupName in _mapGroupNames)
                {
                    // Each group has its own label counter.
                    _mapGroupLabels.Add(1);
                }
            }

            // Keep track of the label numbers.
            _maxLabel = 1;

            bool success;

            int layerNum = 0;
            int layerCount = SelectedLayers.Count;
            foreach (MapLayer selectedLayer in SelectedLayers)
            {
                // Stop if the user cancelled the process.
                if (_dockPane.SearchCancelled)
                    break;

                // Get the layer name.
                string mapNodeGroup = selectedLayer.NodeGroup;
                string mapNodeLayer = selectedLayer.NodeLayer;

                _dockPane.ProgressUpdate("Processing '" + mapNodeGroup + " - " + mapNodeLayer + "'...", stepNum, 0);
                stepNum += 1;

                layerNum += 1;
                FileFunctions.WriteLine(_logFile, "");
                FileFunctions.WriteLine(_logFile, "Starting analysis for '" + selectedLayer.NodeName + "' (" + layerNum + " of " + layerCount + ")");

                // Loop through the map layers, processing each one.
                success = await ProcessMapLayerAsync(selectedLayer, reference, siteName, shortRef, subref, radius, areaMeasureUnit, keepSelectedLayers, addSelectedLayersOption, overwriteLabelOption, combinedSitesTableOption);

                // Keep track of any errors.
                if (!success)
                    _searchErrors = true;
            }

            // Increment the progress value to the last step.
            _dockPane.ProgressUpdate("Cleaning up...", stepNum, 0);

            // If there were errors then exit before cleaning up.
            if (_searchErrors)
                return false;

            // Clean up after the search.
            await CleanUpSearchAsync(addSelectedLayersOption);

            // If the process was cancelled when exit.
            if (_dockPane.SearchCancelled)
                return false;

            // Zoom to the buffer layer extent.
            if (bufferSize != "0" && _keepBuffer)
                await _mapFunctions.ZoomToLayerAsync(_bufferLayerName, 1.05);

            return true;
        }

        /// <summary>
        /// Indicate that the search process has stopped (either
        /// successfully or otherwise).
        /// </summary>
        /// <param name="searchRef"></param>
        /// <param name="success"></param>
        private void StopSearch(string searchRef, string message, string image)
        {
            FileFunctions.WriteLine(_logFile, "---------------------------------------------------------------------------");
            FileFunctions.WriteLine(_logFile, string.Format(message, searchRef));
            FileFunctions.WriteLine(_logFile, "---------------------------------------------------------------------------");

            // Resume the map redrawing.
            _mapFunctions.PauseDrawing(false);

            // Indicate search has finished.
            _dockPane.SearchRunning = false;
            _dockPane.ProgressUpdate(null, -1, -1);

            string imageSource = string.Format("pack://application:,,,/DataSelector;component/Images/{0}32.png", image);

            // Notify user of completion.
            Notification notification = new()
            {
                Title = "Data Searches",
                Severity = Notification.SeverityLevel.High,
                Message = string.Format(message, searchRef),
                ImageSource = new BitmapImage(new Uri(imageSource)) as ImageSource
            };
            FrameworkApplication.AddNotification(notification);

            // Open the log file (if required).
            if (OpenLogFile || _searchErrors)
                Process.Start("notepad.exe", _logFile);
        }

        /// <summary>
        /// Clean up after the search has completed (successfully or not).
        /// </summary>
        /// <param name="addSelectedLayersOption"></param>
        /// <returns></returns>
        private async Task CleanUpSearchAsync(AddSelectedLayersOptions addSelectedLayersOption)
        {
            FileFunctions.WriteLine(_logFile, "");

            // Do we want to keep the buffer layer? If not, remove it.
            if (_keepBuffer)
            {
                // Add the output buffer layer to the map.
                if (addSelectedLayersOption != AddSelectedLayersOptions.No)
                {
                    // Set the buffer layer symbology to use.
                    string symbologyFile = _layerFolder + "\\" + _bufferLayerFile;

                    if (!await SetLayerInMapAsync(_bufferLayerName, symbologyFile, 0))
                    {
                        //MessageBox.Show("Error setting buffer layer in the map.");
                        FileFunctions.WriteLine(_logFile, "Error setting buffer layer in the map");
                        _searchErrors = true;
                    }

                    FileFunctions.WriteLine(_logFile, "Buffer layer added to display");
                }
                else
                {
                    // Remove the buffer layer from the map.
                    await _mapFunctions.RemoveLayerAsync(_bufferLayerName);
                }
            }
            else
            {
                try
                {
                    // Remove the buffer layer from the map.
                    await _mapFunctions.RemoveLayerAsync(_bufferLayerName);

                    // Delete the buffer feature class.
                    await ArcGISFunctions.DeleteFeatureClassAsync(_bufferOutputFile);

                    FileFunctions.WriteLine(_logFile, "Buffer layer deleted");
                }
                catch
                {
                    //MessageBox.Show("Error deleting the buffer layer.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    FileFunctions.WriteLine(_logFile, "Error deleting the buffer layer");
                    _searchErrors = true;
                }
            }

            // Save the selected feature if required.
            if (_keepSearchFeature)
            {
                // Add the search layer to the map.
                if (addSelectedLayersOption != AddSelectedLayersOptions.No)
                {
                    // Set the search layer symbology to use.
                    string searchlayerFile = _searchSymbologyBase + _searchLayerExtension + ".lyrx";
                    string symbologyFile = _layerFolder + "\\" + searchlayerFile;

                    if (!await SetLayerInMapAsync(_searchLayerName, symbologyFile, 0))
                    {
                        //MessageBox.Show("Error setting search feature layer in the map.");
                        FileFunctions.WriteLine(_logFile, "Error setting search feature layer in the map");
                        _searchErrors = true;
                    }

                    FileFunctions.WriteLine(_logFile, "Search feature layer added to display");
                }
                else
                {
                    // Remove the search feature layer from the map.
                    await _mapFunctions.RemoveLayerAsync(_searchLayerName);
                }
            }
            else
            {
                try
                {
                    // Remove the search feature layer from the map.
                    await _mapFunctions.RemoveLayerAsync(_searchLayerName);

                    // Delete the search feature class.
                    string searchOutputFile = _outputFolder + "\\" + _searchLayerName + ".shp";
                    await ArcGISFunctions.DeleteFeatureClassAsync(searchOutputFile);

                    FileFunctions.WriteLine(_logFile, "Search feature layer deleted");
                }
                catch
                {
                    //MessageBox.Show("Error deleting the search feature layer.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    FileFunctions.WriteLine(_logFile, "Error deleting the search feature layer");
                    _searchErrors = true;
                }
            }

            // Remove all temporary feature classes and tables.
            await _mapFunctions.RemoveLayerAsync(_tempMasterLayerName);
            await _mapFunctions.RemoveLayerAsync(_tempFCLayerName);
            await _mapFunctions.RemoveLayerAsync(_tempFCPointsLayerName);
            await _mapFunctions.RemoveLayerAsync(_tempSearchPointsLayerName);
            await _mapFunctions.RemoveTableAsync(_tempTableLayerName);

            // Delete the temporary feature classes and tables.
            if (await ArcGISFunctions.FeatureClassExistsAsync(_tempMasterOutputFile))
                await ArcGISFunctions.DeleteGeodatabaseFCAsync(_tempGDBName, _tempMasterLayerName);

            if (await ArcGISFunctions.FeatureClassExistsAsync(_tempFCOutputFile))
                await ArcGISFunctions.DeleteGeodatabaseFCAsync(_tempGDBName, _tempFCLayerName);

            if (await ArcGISFunctions.FeatureClassExistsAsync(_tempFCPointsOutputFile))
                await ArcGISFunctions.DeleteGeodatabaseFCAsync(_tempGDBName, _tempFCPointsLayerName);

            if (await ArcGISFunctions.FeatureClassExistsAsync(_tempSearchPointsOutputFile))
                await ArcGISFunctions.DeleteGeodatabaseFCAsync(_tempGDBName, _tempSearchPointsLayerName);

            if (await ArcGISFunctions.TableExistsAsync(_tempTableOutputFile))
                await ArcGISFunctions.DeleteGeodatabaseTableAsync(_tempGDBName, _tempTableLayerName);

            // Clear the search features selection.
            await _mapFunctions.ClearLayerSelectionAsync(_inputLayerName);

            // Remove the group layer from the map if it is empty.
            await _mapFunctions.RemoveGroupLayerAsync(_groupLayerName);
        }

        /// <summary>
        /// Create the output folders if required.
        /// </summary>
        /// <param name="saveRootDir"></param>
        /// <param name="saveFolder"></param>
        /// <param name="gisFolder"></param>
        /// <returns></returns>
        private string CreateOutputFolders(string saveRootDir, string saveFolder, string gisFolder)
        {
            // Create root folder if required.
            if (!FileFunctions.DirExists(saveRootDir))
            {
                try
                {
                    Directory.CreateDirectory(saveRootDir);
                }
                catch (Exception ex)
                {
                    FileFunctions.WriteLine(_logFile, "Cannot create directory '" + saveRootDir + "'. System error: " + ex.Message);
                    return null;
                }
            }

            // Create save sub-folder if required.
            if (!string.IsNullOrEmpty(saveFolder))
                saveFolder = saveRootDir + @"\" + saveFolder;
            else
                saveFolder = saveRootDir;
            if (!FileFunctions.DirExists(saveFolder))
            {
                try
                {
                    Directory.CreateDirectory(saveFolder);
                }
                catch (Exception ex)
                {
                    FileFunctions.WriteLine(_logFile, "Cannot create directory '" + saveFolder + "'. System error: " + ex.Message);
                    return null;
                }
            }

            // Create gis sub-folder if required.
            if (!string.IsNullOrEmpty(gisFolder))
                gisFolder = saveFolder + @"\" + gisFolder;
            else
                gisFolder = saveFolder;

            if (!FileFunctions.DirExists(gisFolder))
            {
                try
                {
                    Directory.CreateDirectory(gisFolder);
                }
                catch (Exception ex)
                {
                    FileFunctions.WriteLine(_logFile, "Cannot create directory '" + gisFolder + "'. System error: " + ex.Message);
                    return null;
                }
            }

            return gisFolder;
        }

        /// <summary>
        /// Create an empty combined sites table (with just the column headers).
        /// </summary>
        /// <param name="combinedSitesTable"></param>
        /// <param name="combinedSitesTableOption"></param>
        /// <returns></returns>
        private bool CreateCombinedSitesTable(string combinedSitesTable, CombinedSitesTableOptions combinedSitesTableOption)
        {
            // Start the table if overwrite has been selected, or if the table doesn't exist (and append has been selected).
            if (combinedSitesTableOption == CombinedSitesTableOptions.Overwrite ||
                combinedSitesTableOption == CombinedSitesTableOptions.Append && !FileFunctions.FileExists(combinedSitesTable))
            {
                if (!FileFunctions.WriteEmptyTextFile(combinedSitesTable, _combinedSitesColumnList))
                {
                    //MessageBox.Show("Error writing to combined sites table.");
                    FileFunctions.WriteLine(_logFile, "Error writing to combined sites table");
                    _searchErrors = true;

                    return false;
                }

                FileFunctions.WriteLine(_logFile, "Combined sites table started");
            }

            return true;
        }

        /// <summary>
        /// Find the search reference features in the search layers.
        /// </summary>
        /// <param name="reference"></param>
        /// <returns>Name of the target layer.</returns>
        private async Task<bool> FindSearchFeaturesAsync(string searchClause)
        {
            // Loop through all base layer and extension combinations.
            foreach (string searchLayerExtension in _searchLayerExtensions)
            {
                string searchLayer = _searchLayerBase + searchLayerExtension;

                // Find the feature layer by name if it exists. Only search existing layers.
                FeatureLayer featureLayer = _mapFunctions.FindLayer(searchLayer);

                if (featureLayer != null)
                {
                    // Count the required features in the layer.
                    long featureCount = await ArcGISFunctions.CountFeaturesAsync(featureLayer, searchClause);

                    if (featureCount > 0)
                        return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Count the search reference features in the search layers.
        /// </summary>
        /// <param name="reference"></param>
        /// <returns>Name of the target layer.</returns>
        private async Task<long> CountSearchFeaturesAsync(string searchClause)
        {
            // Find the search feature.
            int featureLayerCount = 0;
            long totalFeatureCount = 0;

            // Loop through all base layer and extension combinations.
            foreach (string searchLayerExtension in _searchLayerExtensions)
            {
                string searchLayer = _searchLayerBase + searchLayerExtension;

                // Find the feature layer by name if it exists. Only search existing layers.
                FeatureLayer featureLayer = _mapFunctions.FindLayer(searchLayer);

                if (featureLayer != null)
                {
                    // Count the required features in the layer.
                    long featureCount = await ArcGISFunctions.CountFeaturesAsync(featureLayer, searchClause);

                    if (featureCount == 0)
                        FileFunctions.WriteLine(_logFile, "No features found in " + searchLayer);

                    if (featureCount > 0)
                    {
                        FileFunctions.WriteLine(_logFile, featureCount.ToString() + " feature(s) found in " + searchLayer);
                        if (featureLayerCount == 0)
                        {
                            // Save the layer name and extension where the feature(s) were found.
                            _inputLayerName = searchLayer;
                            _searchLayerExtension = searchLayerExtension;
                        }
                        totalFeatureCount += featureCount;
                        featureLayerCount++;
                    }
                }
            }

            // If no features found in any layer.
            if (featureLayerCount == 0)
            {
                //MessageBox.Show("No features found in any of the search layers; Process aborted.");
                FileFunctions.WriteLine(_logFile, "No features found in any of the search layers");

                return 0;
            }

            // If features found in more than one layer.
            if (featureLayerCount > 1)
            {
                //MessageBox.Show(totalFeatureCount.ToString() + " features found in different search layers; Process aborted.");
                FileFunctions.WriteLine(_logFile, totalFeatureCount.ToString() + " features found in different search layers");

                return 0;
            }

            // If multiple features found.
            if (totalFeatureCount > 1)
            {
                // Ask the user if they want to continue
                MessageBoxResult response = MessageBox.Show(totalFeatureCount.ToString() + " features found in " + _inputLayerName + " matching those criteria. Do you wish to continue?", "Data Searches", MessageBoxButton.YesNo);
                if (response == MessageBoxResult.No)
                {
                    FileFunctions.WriteLine(_logFile, totalFeatureCount.ToString() + " features found in the search layers");

                    return 0;
                }
            }

            return totalFeatureCount;
        }

        /// <summary>
        /// Prepare a new temporary GDB to use and check it's empty (in case it
        /// already existed).
        /// </summary>
        /// <returns></returns>
        private async Task<bool> PrepareTemporaryGDBAsync()
        {
            // Set a temporary folder path.
            string tempFolder = Path.GetTempPath();

            // Create the temporary file geodatabase if it doesn't exist.
            _tempGDBName = tempFolder + @"Temp.gdb";
            _tempGDB = null;
            bool tempGDBFound = true;
            if (!FileFunctions.DirExists(_tempGDBName))
            {
                _tempGDB = ArcGISFunctions.CreateFileGeodatabase(_tempGDBName);
                if (_tempGDB == null)
                {
                    //MessageBox.Show("Error creating temporary geodatabase" + _tempGDBName);
                    FileFunctions.WriteLine(_logFile, "Error creating temporary geodatabase " + _tempGDBName);
                    _searchErrors = true;

                    return false;
                }

                tempGDBFound = false;
                FileFunctions.WriteLine(_logFile, "Temporary geodatabase created");
            }

            // Set the temporary layer and file names.
            _tempMasterLayerName = "TempMaster_" + _userID;
            _tempMasterOutputFile = _tempGDBName + @"\" + _tempMasterLayerName;
            _tempFCLayerName = "TempOutput_" + _userID;
            _tempFCOutputFile = _tempGDBName + @"\" + _tempFCLayerName;
            _tempFCPointsLayerName = "TempOutputPoints_" + _userID;
            _tempFCPointsOutputFile = _tempGDBName + @"\" + _tempFCPointsLayerName;
            _tempSearchPointsLayerName = "TempSearchPoints_" + _userID;
            _tempSearchPointsOutputFile = _tempGDBName + @"\" + _tempSearchPointsLayerName;
            _tempTableLayerName = "TempTable_" + _userID;
            _tempTableOutputFile = _tempGDBName + @"\" + _tempTableLayerName;

            // If the GDB already existed clean it up.
            if (tempGDBFound)
            {
                // Delete the temporary master feature class if it still exists.
                await _mapFunctions.RemoveLayerAsync(_tempMasterLayerName);
                if (await ArcGISFunctions.FeatureClassExistsAsync(_tempMasterOutputFile))
                {
                    await ArcGISFunctions.DeleteGeodatabaseFCAsync(_tempGDBName, _tempMasterLayerName);
                    //FileFunctions.WriteLine(_logFile, "Temporary master feature class deleted");
                }

                // Delete the temporary output feature class if it still exists.
                await _mapFunctions.RemoveLayerAsync(_tempFCLayerName);
                if (await ArcGISFunctions.FeatureClassExistsAsync(_tempFCOutputFile))
                {
                    await ArcGISFunctions.DeleteGeodatabaseFCAsync(_tempGDBName, _tempFCLayerName);
                    //FileFunctions.WriteLine(_logFile, "Temporary output feature class deleted");
                }

                // Delete the temporary output points feature class if it still exists.
                await _mapFunctions.RemoveLayerAsync(_tempFCPointsLayerName);
                if (await ArcGISFunctions.FeatureClassExistsAsync(_tempFCPointsOutputFile))
                {
                    await ArcGISFunctions.DeleteGeodatabaseFCAsync(_tempGDBName, _tempFCPointsLayerName);
                    //FileFunctions.WriteLine(_logFile, "Temporary output feature class deleted");
                }

                // Delete the temporary search points feature class if it still exists.
                await _mapFunctions.RemoveLayerAsync(_tempSearchPointsLayerName);
                if (await ArcGISFunctions.FeatureClassExistsAsync(_tempSearchPointsOutputFile))
                {
                    await ArcGISFunctions.DeleteGeodatabaseFCAsync(_tempGDBName, _tempSearchPointsLayerName);
                    //FileFunctions.WriteLine(_logFile, "Temporary output feature class deleted");
                }

                // Delete the temporary output table if it still exists.
                await _mapFunctions.RemoveTableAsync(_tempTableLayerName);
                if (await ArcGISFunctions.TableExistsAsync(_tempTableOutputFile))
                {
                    await ArcGISFunctions.DeleteGeodatabaseTableAsync(_tempGDBName, _tempTableLayerName);
                    //FileFunctions.WriteLine(_logFile, "Temporary output table deleted");
                }
            }

            return true;
        }

        /// <summary>
        /// Save the selected search feature(s) to a new layer.
        /// </summary>
        /// <returns></returns>
        private async Task<bool> SaveSearchFeaturesAsync()
        {
            // Get the full layer path (in case it's nested in one or more groups).
            string inputLayerPath = _mapFunctions.GetLayerPath(_inputLayerName);

            FileFunctions.WriteLine(_logFile, "Saving search feature(s)");

            // Copy the selected feature(s) to an output file.
            if (!await ArcGISFunctions.CopyFeaturesAsync(inputLayerPath, _searchOutputFile, true))
            {
                //MessageBox.Show("Error saving search feature(s)");
                FileFunctions.WriteLine(_logFile, "Error saving search feature(s)");
                _searchErrors = true;

                return false;
            }

            return true;
        }

        /// <summary>
        /// Buffer the selected search feature(s) to a new layer.
        /// </summary>
        /// <param name="bufferSize"></param>
        /// <param name="bufferUnitProcess"></param>
        /// <param name="bufferUnitShort"></param>
        /// <returns></returns>
        private async Task<bool> BufferSearchFeaturesAsync(string bufferSize, string bufferUnitProcess, string bufferUnitShort)
        {
            // Get the full layer path (in case it's nested in one or more groups).
            string searchLayerPath = _mapFunctions.GetLayerPath(_searchLayerName);

            // Create a buffer around the feature and save into a new file.
            string bufferDistance = bufferSize + " " + bufferUnitProcess;

            // Safeguard for zero buffer size; Select a tiny buffer to allow
            // correct legending (expects a polygon).
            if (bufferSize == "0")
                bufferDistance = "0.01 Meters";

            // Check if all fields in the aggregate fields exist. If not, ignore.
            List<string> aggColumns = [.. _bufferFields.Split(';')];
            string dissolveFields = "";
            foreach (string fieldName in aggColumns)
            {
                if (await _mapFunctions.FieldExistsAsync(searchLayerPath, fieldName))
                    dissolveFields = dissolveFields + fieldName + ";";
            }

            FileFunctions.WriteLine(_logFile, "Buffering feature(s) with a distance of " + bufferSize + bufferUnitShort);

            string dissolveOption = "ALL";
            if (!string.IsNullOrEmpty(dissolveFields))
            {
                dissolveFields = dissolveFields.Substring(0, dissolveFields.Length - 1);
                dissolveOption = "LIST";
            }

            if (!await ArcGISFunctions.BufferFeaturesAsync(searchLayerPath, _bufferOutputFile, bufferDistance, "FULL", "ROUND", dissolveOption, dissolveFields, addToMap: true))
            {
                //MessageBox.Show("Error during feature buffering.");
                FileFunctions.WriteLine(_logFile, "Error during feature buffering");
                _searchErrors = true;

                return false;
            }

            return true;
        }

        /// <summary>
        /// Apply any symbology to the new layer and move it to the group layer.
        /// </summary>
        /// <param name="layerName"></param>
        /// <param name="symbologyFile"></param>
        /// <param name="position"></param>
        /// <returns></returns>
        private async Task<bool> SetLayerInMapAsync(string layerName, string symbologyFile, int position = -1)
        {
            // Apply layer symbology.
            if (!string.IsNullOrEmpty(symbologyFile) && symbologyFile.Substring(symbologyFile.Length - 4, 4).Equals("lyrx", StringComparison.OrdinalIgnoreCase))
            {
                if (FileFunctions.FileExists(symbologyFile))
                {
                    if (!await _mapFunctions.ApplySymbologyFromLayerFileAsync(layerName, symbologyFile))
                    {
                        //MessageBox.Show("Error applying symbology to '" + layerName + "'");
                        FileFunctions.WriteLine(_logFile, "Error applying symbology to '" + layerName + "'");
                        _searchErrors = true;

                        return false;
                    }
                }
            }

            // Move layer to the group layer.
            if (!string.IsNullOrEmpty(_groupLayerName))
            {
                if (!await _mapFunctions.MoveToGroupLayerAsync(_mapFunctions.FindLayer(layerName), _groupLayerName, position))
                {
                    //MessageBox.Show("Error moving layer to '" + _groupLayerName + "'");
                    FileFunctions.WriteLine(_logFile, "Error moving layer to '" + _groupLayerName + "'");
                    _searchErrors = true;

                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Process each of the selected map layers.
        /// </summary>
        /// <param name="selectedLayer"></param>
        /// <param name="reference"></param>
        /// <param name="siteName"></param>
        /// <param name="shortRef"></param>
        /// <param name="subref"></param>
        /// <param name="radius"></param>
        /// <param name="areaMeasureUnit"></param>
        /// <param name="keepSelectedLayers"></param>
        /// <param name="addSelectedLayersOption"></param>
        /// <param name="overwriteLabelOption"></param>
        /// <param name="combinedSitesTableOption"></param>
        /// <returns></returns>
        private async Task<bool> ProcessMapLayerAsync(MapLayer selectedLayer, string reference, string siteName,
            string shortRef, string subref, string radius, string areaMeasureUnit,
            bool keepSelectedLayers,
            AddSelectedLayersOptions addSelectedLayersOption,
            OverwriteLabelOptions overwriteLabelOption,
            CombinedSitesTableOptions combinedSitesTableOption)
        {
            // Get the settings relevant for this layer.
            string mapNodeGroup = selectedLayer.NodeGroup;
            //string mapNodeLayer = selectedLayer.NodeLayer;
            string mapLayerName = selectedLayer.LayerName;
            string mapOutputName = selectedLayer.GISOutputName;
            string mapTableOutputName = selectedLayer.TableOutputName;
            string mapColumns = selectedLayer.Columns;
            string mapGroupColumns = selectedLayer.GroupColumns;
            string mapStatsColumns = selectedLayer.StatisticsColumns;
            string mapOrderColumns = selectedLayer.OrderColumns;
            string mapCriteria = selectedLayer.Criteria;

            bool mapIncludeArea = selectedLayer.IncludeArea;
            string mapIncludeNearFields = selectedLayer.IncludeNearFields;
            bool mapIncludeRadius = selectedLayer.IncludeRadius;

            string mapKeyColumn = selectedLayer.KeyColumn;
            string mapFormat = selectedLayer.Format;
            bool mapKeepLayer = selectedLayer.KeepLayer;
            string mapOutputType = selectedLayer.OutputType;

            bool mapDisplayLabels = selectedLayer.DisplayLabels;
            string mapLayerFileName = selectedLayer.LayerFileName;
            bool mapOverwriteLabels = selectedLayer.OverwriteLabels;
            string mapLabelColumn = selectedLayer.LabelColumn;
            string mapLabelClause = selectedLayer.LabelClause;
            string mapMacroName = selectedLayer.MacroName;

            string mapCombinedSitesColumns = selectedLayer.CombinedSitesColumns;
            string mapCombinedSitesGroupColumns = selectedLayer.CombinedSitesGroupColumns;
            string mapCombinedSitesStatsColumns = selectedLayer.CombinedSitesStatisticsColumns;
            string mapCombinedSitesOrderColumns = selectedLayer.CombinedSitesOrderByColumns;

            // Deal with wildcards in the output names.
            mapOutputName = StringFunctions.ReplaceSearchStrings(mapOutputName, reference, siteName, shortRef, subref, radius);
            mapTableOutputName = StringFunctions.ReplaceSearchStrings(mapTableOutputName, reference, siteName, shortRef, subref, radius);

            // Remove any illegal characters from the names.
            mapOutputName = StringFunctions.StripIllegals(mapOutputName, _repChar);
            mapTableOutputName = StringFunctions.StripIllegals(mapTableOutputName, _repChar);

            // Set the statistics columns if they haven't been supplied.
            if (String.IsNullOrEmpty(mapStatsColumns) && !String.IsNullOrEmpty(mapGroupColumns))
                mapStatsColumns = StringFunctions.AlignStatsColumns(mapColumns, mapStatsColumns, mapGroupColumns);
            if (String.IsNullOrEmpty(mapCombinedSitesStatsColumns) && !String.IsNullOrEmpty(mapCombinedSitesGroupColumns))
                mapCombinedSitesStatsColumns = StringFunctions.AlignStatsColumns(mapCombinedSitesColumns, mapCombinedSitesStatsColumns, mapCombinedSitesGroupColumns);

            // Get the full layer path (in case it's nested in one or more groups).
            string mapLayerPath = _mapFunctions.GetLayerPath(mapLayerName);

            // Select by location.
            FileFunctions.WriteLine(_logFile, "Selecting features using selected feature(s) from layer " + _bufferLayerName + " ...");
            if (!await ArcGISFunctions.SelectLayerByLocationAsync(mapLayerPath, _bufferLayerPath, "INTERSECT", "", "NEW_SELECTION"))
            {
                //MessageBox.Show("Error selecting layer " + mapLayerName + " by location.");
                FileFunctions.WriteLine(_logFile, "Error selecting layer " + mapLayerName + " by location");
                _searchErrors = true;

                return false;
            }

            // Find the map layer by name.
            FeatureLayer mapLayer = _mapFunctions.FindLayer(mapLayerName);

            if (mapLayer == null)
                return false;

            // Refine the selection by attributes (if required).
            if (mapLayer.SelectionCount > 0 && !string.IsNullOrEmpty(mapCriteria))
            {
                FileFunctions.WriteLine(_logFile, "Refining selection with criteria " + mapCriteria + " ...");

                if (!await _mapFunctions.SelectLayerByAttributesAsync(mapLayerName, mapCriteria, SelectionCombinationMethod.And))
                {
                    //MessageBox.Show("Error selecting layer " + mapLayerName + " with criteria " + mapCriteria + ". Please check syntax and column names (case sensitive).");
                    FileFunctions.WriteLine(_logFile, "Error refining selection on layer " + mapLayerName + " with criteria " + mapCriteria + ". Please check syntax and column names (case sensitive)");
                    _searchErrors = true;

                    return false;
                }
            }

            // Count the selected features.
            int featureCount = mapLayer.SelectionCount;

            // Write out the results - to a feature class initially. Include distance if required.
            if (featureCount <= 0)
            {
                FileFunctions.WriteLine(_logFile, "No features found");
                return true;
            }

            FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", featureCount) + " feature(s) found");

            // Create the map output depending on the output type required.
            if (!await CreateMapOutputAsync(mapLayerName, mapLayerPath, _bufferLayerPath, mapOutputType))
            {
                MessageBox.Show("Cannot output selection from " + mapLayerName + " to " + _tempMasterOutputFile + ".");
                FileFunctions.WriteLine(_logFile, "Cannot output selection from " + mapLayerName + " to " + _tempMasterOutputFile);

                return false;
            }

            // Add map labels to the output if required.
            if (addSelectedLayersOption == AddSelectedLayersOptions.WithLabels && !string.IsNullOrEmpty(mapLabelColumn))
            {
                if (!await AddMapLabelsAsync(overwriteLabelOption, mapOverwriteLabels, mapLabelColumn, mapKeyColumn, mapNodeGroup))
                {
                    //MessageBox.Show("Error adding map labels to " + mapLabelColumn + " in " + _tempMasterOutputFile + ".");
                    FileFunctions.WriteLine(_logFile, "Error adding map labels to " + mapLabelColumn + " in " + _tempMasterOutputFile);
                    _searchErrors = true;

                    return false;
                }
            }

            // Create relevant output names.
            string mapOutputFile = _outputFolder + @"\" + mapOutputName; // Output shapefile / feature class name. Note no extension to allow write to GDB.
            string mapTableOutputFile = _outputFolder + @"\" + mapTableOutputName + "." + mapFormat.ToLower(); // Output table name.

            // Include headers for CSV files.
            bool includeHeaders = false;
            if (mapFormat.Equals("csv", StringComparison.OrdinalIgnoreCase))
                includeHeaders = true;

            // Only include radius if requested.
            string radiusText = "none";
            if (mapIncludeRadius)
                radiusText = radius;

            string areaUnit = "";
            if (mapIncludeArea)
                areaUnit = areaMeasureUnit;

            // Export results to table if required.
            if (!string.IsNullOrEmpty(mapFormat) && !string.IsNullOrEmpty(mapColumns))
            {
                FileFunctions.WriteLine(_logFile, "Extracting summary information ...");

                int intLineCount = await ExportSelectionAsync(mapTableOutputFile, mapFormat.ToLower(), mapColumns, mapGroupColumns, mapStatsColumns, mapOrderColumns,
                    includeHeaders, false, areaUnit, mapIncludeNearFields, radiusText);
                if (intLineCount <= 0)
                {
                    //MessageBox.Show("Error extracting summary from " + _tempMasterOutputFile + ".");
                    FileFunctions.WriteLine(_logFile, "Error extracting summary from " + _tempMasterOutputFile);
                    _searchErrors = true;

                    return false;
                }

                FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", intLineCount) + " record(s) exported");
            }

            // If selected layers are to be kept, and this layer is to be kept,
            // copy to a permanent layer.
            if ((keepSelectedLayers) && (mapKeepLayer))
            {
                if (!await KeepLayerAsync(mapOutputName, mapOutputFile, addSelectedLayersOption, mapLayerFileName, mapDisplayLabels, mapLabelClause, mapLabelColumn))
                {
                    _searchErrors = true;

                    return false;
                }
            }

            // Add to combined sites table if required.
            if (!string.IsNullOrEmpty(mapCombinedSitesColumns) && combinedSitesTableOption != CombinedSitesTableOptions.None)
            {
                FileFunctions.WriteLine(_logFile, "Extracting summary output for combined sites table ...");

                int intLineCount = await ExportSelectionAsync(_combinedSitesOutputFile, _combinedSitesTableFormat, mapCombinedSitesColumns, mapCombinedSitesGroupColumns,
                    mapCombinedSitesStatsColumns, mapCombinedSitesOrderColumns,
                    false, true, areaUnit, mapIncludeNearFields, radiusText);

                if (intLineCount < 0)
                {
                    //MessageBox.Show("Error extracting summary for combined sites table from " + _tempMasterOutputFile + ".");
                    FileFunctions.WriteLine(_logFile, "Error extracting summary for combined sites table from " + _tempMasterOutputFile);
                    _searchErrors = true;

                    return false;
                }

                FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", intLineCount) + " row(s) added to combined sites table");
            }

            // Cleanup the temporary master layer.
            //await _mapFunctions.RemoveLayerAsync(_tempMasterLayerName);
            //await ArcGISFunctions.DeleteFeatureClassAsync(_tempMasterOutputFile);

            // Clear the selection in the input layer.
            await _mapFunctions.ClearLayerSelectionAsync(mapLayerName);

            FileFunctions.WriteLine(_logFile, "Analysis complete");

            // Trigger the macro if one exists
            if (!string.IsNullOrEmpty(mapMacroName))
            {
                FileFunctions.WriteLine(_logFile, "Executing vbscript macro ...");

                if (!StartProcess(mapMacroName, mapTableOutputName, mapFormat))
                {
                    //MessageBox.Show("Error executing vbscript macro " + mapMacroName + ".");
                    FileFunctions.WriteLine(_logFile, "Error executing vbscript macro " + mapMacroName);
                    _searchErrors = true;

                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Create the required output type from the current layer.
        /// </summary>
        /// <param name="mapLayerName"></param>
        /// <param name="mapLayerPath"></param>
        /// <param name="bufferLayerPath"></param>
        /// <param name="mapOutputType"></param>
        /// <returns></returns>
        private async Task<bool> CreateMapOutputAsync(string mapLayerName, string mapLayerPath, string bufferLayerPath, string mapOutputType)
        {
            // Get the input feature class type.
            string mapLayerFCType = _mapFunctions.GetFeatureClassType(mapLayerName);
            if (mapLayerFCType == null)
                return false;

            // Get the buffer feature class type.
            string bufferFCType = _mapFunctions.GetFeatureClassType(_bufferLayerName);
            if (bufferFCType == null)
                return false;

            // If the input layer should be clipped to the buffer layer, do so now.
            if (mapOutputType.Equals("CLIP", StringComparison.OrdinalIgnoreCase))
            {
                if (mapLayerFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase) && bufferFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase) ||
                    mapLayerFCType.Equals("line", StringComparison.OrdinalIgnoreCase) &&
                    (bufferFCType.Equals("line", StringComparison.OrdinalIgnoreCase) || bufferFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase)))
                {
                    // Clip
                    FileFunctions.WriteLine(_logFile, "Clipping selected features ...");
                    return await ArcGISFunctions.ClipFeaturesAsync(mapLayerPath, bufferLayerPath, _tempMasterOutputFile, true);
                }
                else
                {
                    // Copy
                    FileFunctions.WriteLine(_logFile, "Copying selected features ...");
                    return await ArcGISFunctions.CopyFeaturesAsync(mapLayerPath, _tempMasterOutputFile, true);
                }
            }
            // If the buffer layer should be clipped to the input layer, do so now.
            else if (mapOutputType.Equals("OVERLAY", StringComparison.OrdinalIgnoreCase))
            {
                if (bufferFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase) && mapLayerFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase) ||
                    bufferFCType.Equals("line", StringComparison.OrdinalIgnoreCase) &&
                    (mapLayerFCType.Equals("line", StringComparison.OrdinalIgnoreCase) || mapLayerFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase)))
                {
                    // Clip
                    FileFunctions.WriteLine(_logFile, "Overlaying selected features ...");
                    return await ArcGISFunctions.ClipFeaturesAsync(bufferLayerPath, mapLayerPath, _tempMasterOutputFile, true);
                }
                else
                {
                    // Select from the buffer layer.
                    FileFunctions.WriteLine(_logFile, "Selecting features  ...");
                    await ArcGISFunctions.SelectLayerByLocationAsync(bufferLayerPath, mapLayerPath);

                    // Find the buffer layer by name.
                    FeatureLayer bufferLayer = _mapFunctions.FindLayer(_bufferLayerName);

                    if (bufferLayer == null)
                        return false;

                    // Count the selected features.
                    int featureCount = bufferLayer.SelectionCount;
                    if (featureCount > 0)
                    {
                        FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", featureCount) + " feature(s) found");

                        // Copy the selection from the buffer layer.
                        FileFunctions.WriteLine(_logFile, "Copying selected features ... ");
                        if (!await ArcGISFunctions.CopyFeaturesAsync(bufferLayerPath, _tempMasterOutputFile, true))
                            return false;
                    }
                    else
                    {
                        FileFunctions.WriteLine(_logFile, "No features selected");

                        return true;
                    }

                    // Clear the buffer layer selection.
                    await _mapFunctions.ClearLayerSelectionAsync(_bufferLayerName);

                    return true;
                }
            }
            // If the input layer should be intersected with the buffer layer, do so now.
            else if (mapOutputType.Equals("INTERSECT", StringComparison.OrdinalIgnoreCase))
            {
                if (mapLayerFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase) && bufferFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase) ||
                    mapLayerFCType.Equals("line", StringComparison.OrdinalIgnoreCase) &&
                    (bufferFCType.Equals("line", StringComparison.OrdinalIgnoreCase) || bufferFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase)))
                {
                    string[] features = ["'" + mapLayerPath + "' #", "'" + bufferLayerPath + "' #"];
                    string inFeatures = string.Join(";", features);

                    // Intersect
                    FileFunctions.WriteLine(_logFile, "Intersecting selected features ...");
                    return await ArcGISFunctions.IntersectFeaturesAsync(inFeatures, _tempMasterOutputFile, addToMap: true); // Selected features in input, buffer FC, output.
                }
                else
                {
                    // Copy
                    FileFunctions.WriteLine(_logFile, "Copying selected features ...");
                    return await ArcGISFunctions.CopyFeaturesAsync(mapLayerPath, _tempMasterOutputFile, true);
                }
            }
            // Otherwise do a straight copy of the input layer.
            else
            {
                // Copy
                FileFunctions.WriteLine(_logFile, "Copying selected features ...");
                return await ArcGISFunctions.CopyFeaturesAsync(mapLayerPath, _tempMasterOutputFile, true);
            }
        }

        /// <summary>
        /// Add the map labels to the current layer.
        /// </summary>
        /// <param name="overwriteLabelOption"></param>
        /// <param name="overwriteLabels"></param>
        /// <param name="mapLabelColumn"></param>
        /// <param name="mapKeyColumn"></param>
        /// <param name="mapGroupName"></param>
        /// <returns></returns>
        private async Task<bool> AddMapLabelsAsync(OverwriteLabelOptions overwriteLabelOption, bool overwriteLabels,
            string mapLabelColumn, string mapKeyColumn, string mapGroupName)
        {
            bool newLabelField = false;
            // Does the map label field already exist? If not, add it.
            if (!await _mapFunctions.FieldExistsAsync(_tempMasterLayerName, mapLabelColumn))
            {
                if (!await ArcGISFunctions.AddFieldAsync(_tempMasterOutputFile, mapLabelColumn, "LONG"))
                {
                    //MessageBox.Show("Error adding map label field '" + mapLabelColumn + "' to " + _tempMasterOutputFile + ".");
                    FileFunctions.WriteLine(_logFile, "Error adding map label field '" + mapLabelColumn + "' to " + _tempMasterOutputFile);
                    _searchErrors = true;

                    return false;
                }

                newLabelField = true;
            }

            // Either we have a new label field, or we want to overwrite the labels and are allowed to.
            if (newLabelField ||
                overwriteLabelOption != OverwriteLabelOptions.No &&
                overwriteLabels)
            {
                // Add labels as required.
                if (!await CreateMapLabelsAsync(overwriteLabelOption, mapLabelColumn, mapKeyColumn, mapGroupName))
                {
                    //MessageBox.Show("Error setting map labels to " + mapLabelColumn + " in " + _tempMasterOutputFile + ".");
                    FileFunctions.WriteLine(_logFile, "Error setting map labels to " + mapLabelColumn + " in " + _tempMasterOutputFile);
                    _searchErrors = true;

                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Create map labels for the current layer.
        /// </summary>
        /// <param name="overwriteLabelOption"></param>
        /// <param name="mapLabelColumn"></param>
        /// <param name="mapKeyColumn"></param>
        /// <param name="mapGroupName"></param>
        /// <returns></returns>
        private async Task<bool> CreateMapLabelsAsync(OverwriteLabelOptions overwriteLabelOption, string mapLabelColumn, string mapKeyColumn,
            string mapGroupName)
        {
            FileFunctions.WriteLine(_logFile, "Adding map labels ...");

            // Add relevant labels.
            if (overwriteLabelOption == OverwriteLabelOptions.ResetByLayer) // Reset each layer to 1.
            {
                FileFunctions.WriteLine(_logFile, "Resetting label counter ...");

                if (await _mapFunctions.AddIncrementalNumbersAsync(_tempMasterOutputFile, _tempMasterLayerName, mapLabelColumn, mapKeyColumn, 1) < 0)
                {
                    //MessageBox.Show("Error calculating map label field '" + mapLabelColumn + "' in " + _tempMasterOutputFile + ".");
                    FileFunctions.WriteLine(_logFile, "Error calculating map label field '" + mapLabelColumn + "' in " + _tempMasterOutputFile);
                    _searchErrors = true;

                    return false;
                }
            }
            else if (overwriteLabelOption == OverwriteLabelOptions.ResetByGroup && !string.IsNullOrEmpty(mapGroupName))
            {
                // Increment within but reset between groups. Note all group labels are already initialised as 1.
                // Only triggered if a group name has been found.
                int groupIndex = _mapGroupNames.IndexOf(mapGroupName);
                int groupLabel = _mapGroupLabels[groupIndex];

                groupLabel = await _mapFunctions.AddIncrementalNumbersAsync(_tempMasterOutputFile, _tempMasterLayerName, mapLabelColumn, mapKeyColumn, groupLabel);

                // Increment the new group label.
                groupLabel++;

                // Store the new group label.
                _mapGroupLabels[groupIndex] = groupLabel;
            }
            else
            {
                // There is no group or groups are ignored, or we are not resetting. Use the existing max label number.
                int startLabel = _maxLabel;

                _maxLabel = await _mapFunctions.AddIncrementalNumbersAsync(_tempMasterOutputFile, _tempMasterLayerName, mapLabelColumn, mapKeyColumn, startLabel);

                // Increment the new max label.
                _maxLabel++;
            }

            return true;
        }

        /// <summary>
        /// Export the selected features from the current layer to a text file.
        /// </summary>
        /// <param name="outputTableName"></param>
        /// <param name="outputFormat"></param>
        /// <param name="mapColumns"></param>
        /// <param name="mapGroupColumns"></param>
        /// <param name="mapStatsColumns"></param>
        /// <param name="mapOrderColumns"></param>
        /// <param name="includeHeaders"></param>
        /// <param name="append"></param>
        /// <param name="areaUnit"></param>
        /// <param name="includeNearFields"></param>
        /// <param name="radiusText"></param>
        /// <returns></returns>
        private async Task<int> ExportSelectionAsync(string outputTableName, string outputFormat,
            string mapColumns, string mapGroupColumns, string mapStatsColumns, string mapOrderColumns,
            bool includeHeaders, bool append, string areaUnit, string includeNearFields, string radiusText)
        {
            int intLineCount;

            // Only export if the user has specified columns.
            if (string.IsNullOrEmpty(mapColumns))
                return -1;

            // Check the input feature layer exists.
            FeatureLayer inputFeaturelayer = _mapFunctions.FindLayer(_tempMasterLayerName);
            if (inputFeaturelayer == null)
                return -1;

            // Get the input feature class type.
            string inputFeatureType = _mapFunctions.GetFeatureClassType(inputFeaturelayer);
            if (inputFeatureType == null)
                return -1;

            // Calculate the area field if required.
            string areaColumnName = "";
            if (!string.IsNullOrEmpty(areaUnit) && inputFeatureType.Equals("polygon", StringComparison.OrdinalIgnoreCase))
            {
                areaColumnName = "Area" + areaUnit;
                // Does the area field already exist? If not, add it.
                if (!await _mapFunctions.FieldExistsAsync(_tempMasterLayerName, areaColumnName))
                {
                    if (!await ArcGISFunctions.AddFieldAsync(_tempMasterOutputFile, areaColumnName, "DOUBLE", 20))
                    {
                        //MessageBox.Show("Error adding area field to " + _tempMasterOutputFile + ".");
                        FileFunctions.WriteLine(_logFile, "Error adding area field to " + _tempMasterOutputFile);
                        _searchErrors = true;

                        return -1;
                    }

                    string geometryProperty = areaColumnName + " AREA";
                    if (areaUnit.Equals("ha", StringComparison.OrdinalIgnoreCase))
                    {
                        areaUnit = "HECTARES";
                    }
                    else if (areaUnit.Equals("m2", StringComparison.OrdinalIgnoreCase))
                    {
                        areaUnit = "SQUARE_METERS";
                    }
                    else if (areaUnit.Equals("km2", StringComparison.OrdinalIgnoreCase))
                    {
                        areaUnit = "SQUARE_KILOMETERS";
                    }

                    // Calculate the area field.
                    if (!await ArcGISFunctions.CalculateGeometryAsync(_tempMasterOutputFile, geometryProperty, "", areaUnit))
                    {
                        //MessageBox.Show("Error calculating area field in " + _tempMasterOutputFile + ".");
                        FileFunctions.WriteLine(_logFile, "Error calculating area field in " + _tempMasterOutputFile);
                        _searchErrors = true;

                        return -1;
                    }
                }
            }

            // Include radius if requested
            if (radiusText != "none")
            {
                FileFunctions.WriteLine(_logFile, "Including radius column ...");

                // Does the radius field already exist? If not, add it.
                if (!await _mapFunctions.FieldExistsAsync(_tempMasterLayerName, "Radius"))
                {
                    if (!await ArcGISFunctions.AddFieldAsync(_tempMasterOutputFile, "Radius", "TEXT", fieldLength: 25))
                    {
                        //MessageBox.Show("Error adding radius field to " + _tempMasterOutputFile + ".");
                        FileFunctions.WriteLine(_logFile, "Error adding radius field to " + _tempMasterOutputFile);
                        _searchErrors = true;

                        return -1;
                    }
                }

                // Calculate the radius field.
                if (!await ArcGISFunctions.CalculateFieldAsync(_tempMasterOutputFile, "Radius", '"' + radiusText + '"'))
                {
                    //MessageBox.Show("Error calculating radius field in " + _tempMasterOutputFile + ".");
                    FileFunctions.WriteLine(_logFile, "Error calculating radius field in " + _tempMasterOutputFile);
                    _searchErrors = true;

                    return -1;
                }
            }

            // Copy the input features.
            if (!await ArcGISFunctions.CopyFeaturesAsync(_tempMasterOutputFile, _tempFCOutputFile, true))
            {
                //MessageBox.Show("Error copying output file to " + _tempFCOutputFile + ".");
                FileFunctions.WriteLine(_logFile, "Error copying output file to " + _tempFCOutputFile);
                _searchErrors = true;

                return -1;
            }

            //-------------------------------------------------------------
            // After this the input to the remainder of the function
            // should be reading from _tempFCOutputFile (_tempFCLayerName).
            //-------------------------------------------------------------

            // Calculate the boundary distance and bearing if required.
            if (includeNearFields.Equals("BOUNDARY", StringComparison.OrdinalIgnoreCase))
            {
                // Calculate the distance and additional proximity fields.
                if (!await ArcGISFunctions.NearAnalysisAsync(_tempFCOutputFile, _searchLayerName,
                    radiusText, "LOCATION", "ANGLE", "PLANAR", null, "METERS"))
                {
                    //MessageBox.Show("Error calculating nearest distance from " + _tempFCPointsOutputFile + " to " + _tempSearchPointsOutputFile + ".");
                    FileFunctions.WriteLine(_logFile, "Error calculating nearest distance from " + _tempFCOutputFile + " to " + _searchLayerName);
                    _searchErrors = true;

                    return -1;
                }
            }
            // Calculate the centroid distance and bearing if required.
            else if (includeNearFields.Equals("CENTROID", StringComparison.OrdinalIgnoreCase))
            {
                // Convert the output features to points.
                if (!await ArcGISFunctions.FeatureToPointAsync(_tempFCOutputFile, _tempFCPointsOutputFile,
                    "CENTROID", addToMap: false))
                {
                    //MessageBox.Show("Error converting " + _tempFCOutputFile + " features to points into " + _tempFCPointsOutputFile + ".");
                    FileFunctions.WriteLine(_logFile, "Error converting " + _tempFCOutputFile + " features to points into " + _tempFCPointsOutputFile);
                    _searchErrors = true;

                    return -1;
                }

                // Convert the search features to points.
                if (!await ArcGISFunctions.FeatureToPointAsync(_searchLayerName, _tempSearchPointsOutputFile,
                    "CENTROID", addToMap: false))
                {
                    //MessageBox.Show("Error converting " + _searchLayerName + " features to points into " + _tempSearchPointsOutputFile + ".");
                    FileFunctions.WriteLine(_logFile, "Error converting " + _searchLayerName + " features to points into " + _tempSearchPointsOutputFile);
                    _searchErrors = true;

                    return -1;
                }

                // Calculate the distance and additional proximity fields.
                if (!await ArcGISFunctions.NearAnalysisAsync(_tempFCPointsOutputFile, _tempSearchPointsOutputFile,
                    radiusText, "LOCATION", "ANGLE", "PLANAR", null, "METERS"))
                {
                    //MessageBox.Show("Error calculating nearest distance from " + _tempFCPointsOutputFile + " to " + _tempSearchPointsOutputFile + ".");
                    FileFunctions.WriteLine(_logFile, "Error calculating nearest distance from " + _tempFCPointsOutputFile + " to " + _tempSearchPointsOutputFile);
                    _searchErrors = true;

                    return -1;
                }

                string joinFields = "NEAR_DIST;NEAR_ANGLE";

                // Join the distance and addition proximity fields to the output feature layer.
                if (!await ArcGISFunctions.JoinFieldsAsync(_tempFCLayerName, "OBJECTID", _tempFCPointsOutputFile, "ORIG_FID",
                    joinFields, addToMap: true))
                {
                    //MessageBox.Show("Error joining fields to " + _tempFCLayerName + " from " + _tempFCPointsOutputFile + ".");
                    FileFunctions.WriteLine(_logFile, "Error joining fields to " + _tempFCLayerName + " from " + _tempFCPointsOutputFile);
                    _searchErrors = true;

                    return -1;
                }
            }

            // Check the output feature layer exists.
            FeatureLayer outputFeatureLayer = _mapFunctions.FindLayer(_tempFCLayerName);
            if (outputFeatureLayer == null)
                return -1;

            // Check all the requested group by fields exist.
            // Only pass those that do.
            if (!string.IsNullOrEmpty(mapGroupColumns))
            {
                List<string> groupColumnList = [.. mapGroupColumns.Split(';')];
                mapGroupColumns = "";
                foreach (string groupColumn in groupColumnList)
                {
                    string columnName = groupColumn.Trim();

                    if (await _mapFunctions.FieldExistsAsync(_tempFCLayerName, columnName))
                        mapGroupColumns = mapGroupColumns + columnName + ";";
                }
                if (!string.IsNullOrEmpty(mapGroupColumns))
                    mapGroupColumns = mapGroupColumns.Substring(0, mapGroupColumns.Length - 1);
            }

            // Check all the requested statistics fields exist.
            // Only pass those that do.
            if (!string.IsNullOrEmpty(mapStatsColumns))
            {
                List<string> statsColumnList = [.. mapStatsColumns.Split(';')];
                mapStatsColumns = "";
                foreach (string statsColumn in statsColumnList)
                {
                    List<string> statsComponents = [.. statsColumn.Split(' ')];
                    string columnName = statsComponents[0].Trim(); // The field name.

                    if (await _mapFunctions.FieldExistsAsync(_tempFCLayerName, columnName))
                        mapStatsColumns = mapStatsColumns + statsColumn + ";";
                }
                if (!string.IsNullOrEmpty(mapStatsColumns))
                    mapStatsColumns = mapStatsColumns.Substring(0, mapStatsColumns.Length - 1);
            }

            // If we have group columns but no statistics columns, add a dummy column.
            if (string.IsNullOrEmpty(mapStatsColumns) && !string.IsNullOrEmpty(mapGroupColumns))
            {
                string strDummyField = mapGroupColumns.Split(';').ToList()[0];
                mapStatsColumns = strDummyField + " FIRST";
            }

            // Now do the summary statistics as required, or export the layer to table if not.
            if (!string.IsNullOrEmpty(mapStatsColumns))
            {
                FileFunctions.WriteLine(_logFile, "Calculating summary statistics ...");

                string statisticsFields = "";
                if (!string.IsNullOrEmpty(mapStatsColumns))
                    statisticsFields = mapStatsColumns;

                string caseFields = "";
                if (!string.IsNullOrEmpty(mapGroupColumns))
                    caseFields = mapGroupColumns;

                // Add the radius column to the stats columns if it's not already there.
                if (radiusText != "none")
                {
                    if (!statisticsFields.Contains("Radius FIRST", StringComparison.OrdinalIgnoreCase))
                        statisticsFields += ";Radius FIRST";
                }

                // Add the area column to the stats columns if it's not already there.
                if (!string.IsNullOrEmpty(areaColumnName))
                {
                    if (!statisticsFields.Contains(areaColumnName + " FIRST", StringComparison.OrdinalIgnoreCase))
                        statisticsFields += ";" + areaColumnName + " FIRST";
                }

                // Calculate the summary statistics.
                if (!await ArcGISFunctions.CalculateSummaryStatisticsAsync(_tempFCOutputFile, _tempTableOutputFile, statisticsFields, caseFields, addToMap: true))
                {
                    //MessageBox.Show("Error calculating summary statistics for '" + _tempFCOutputFile + "' into " + _tempTableOutputFile + ".");
                    FileFunctions.WriteLine(_logFile, "Error calculating summary statistics for '" + _tempFCOutputFile + "' into " + _tempTableOutputFile);
                    _searchErrors = true;

                    return -1;
                }

                // Get the list of fields for the input table.
                IReadOnlyList<Field> inputFields;
                inputFields = await _mapFunctions.GetTableFieldsAsync(_tempTableLayerName);

                // Check a list of fields is returned.
                if (inputFields == null || inputFields.Count == 0)
                    return -1;

                // Now rename the radius field.
                if (radiusText != "none")
                {
                    string oldFieldName;
                    // Check the radius field by name.
                    try
                    {
                        oldFieldName = inputFields.Where(f => f.Name.Equals("FIRST_Radius", StringComparison.OrdinalIgnoreCase)).First().Name;
                    }
                    catch
                    {
                        // If not found then use the last field.
                        int intNewIndex = inputFields.Count - 1;
                        oldFieldName = inputFields[intNewIndex].Name;
                    }

                    if (!await ArcGISFunctions.RenameFieldAsync(_tempTableOutputFile, oldFieldName, "Radius"))
                    {
                        //MessageBox.Show("Error renaming radius field in " + _tempFCOutputFile + ".");
                        FileFunctions.WriteLine(_logFile, "Error renaming radius field in " + _tempTableLayerName);
                        _searchErrors = true;

                        return -1;
                    }
                }

                // Now rename the area field.
                if (!string.IsNullOrEmpty(areaColumnName))
                {
                    string oldFieldName;
                    // Check the area field by name.
                    try
                    {
                        oldFieldName = inputFields.Where(f => f.Name.Equals("FIRST_" + areaColumnName, StringComparison.OrdinalIgnoreCase)).First().Name;
                    }
                    catch
                    {
                        // If not found then use the last field.
                        int intNewIndex = inputFields.Count - 1;
                        oldFieldName = inputFields[intNewIndex].Name;
                    }

                    if (!await ArcGISFunctions.RenameFieldAsync(_tempTableOutputFile, oldFieldName, areaColumnName))
                    {
                        //MessageBox.Show("Error renaming Area field in " + _tempFCOutputFile + ".");
                        FileFunctions.WriteLine(_logFile, "Error renaming Area field in " + _tempTableLayerName);
                        _searchErrors = true;

                        return -1;
                    }
                }

                // Now export the output table.
                FileFunctions.WriteLine(_logFile, "Exporting to " + outputFormat.ToUpper() + " ...");
                intLineCount = await _mapFunctions.CopyTableToTextFileAsync(_tempTableLayerName, outputTableName, mapColumns, mapOrderColumns, append, includeHeaders);
            }
            else
            {
                // Do straight copy of the feature class.
                FileFunctions.WriteLine(_logFile, "Exporting to " + outputFormat.ToUpper() + " ...");
                intLineCount = await _mapFunctions.CopyFCToTextFileAsync(_tempFCLayerName, outputTableName, mapColumns, mapOrderColumns, append, includeHeaders);
            }

            return intLineCount;
        }

        /// <summary>
        /// Save the selected features from the current layer to a new layer
        /// and add it to the map if required.
        /// </summary>
        /// <param name="layerName"></param>
        /// <param name="outputFile"></param>
        /// <param name="addSelectedLayersOption"></param>
        /// <param name="layerFileName"></param>
        /// <param name="displayLabels"></param>
        /// <param name="labelClause"></param>
        /// <param name="labelColumn"></param>
        /// <returns></returns>
        private async Task<bool> KeepLayerAsync(string layerName, string outputFile, AddSelectedLayersOptions addSelectedLayersOption,
            string layerFileName, bool displayLabels, string labelClause, string labelColumn)
        {
            bool addToMap = addSelectedLayersOption != AddSelectedLayersOptions.No;

            // Copy to a permanent file (note this is not the summarised layer).
            FileFunctions.WriteLine(_logFile, "Copying selected GIS features to " + layerName + ".shp ...");
            await ArcGISFunctions.CopyFeaturesAsync(_tempMasterLayerName, outputFile, addToMap);

            // If the layer is to be added to the map
            if (addToMap)
            {
                FileFunctions.WriteLine(_logFile, "Output " + layerName + " added to display");

                string symbologyFile = null;

                // If there is a layer file to apply.
                if (!string.IsNullOrEmpty(layerFileName))
                {
                    // Set the layer symbology to use.
                    symbologyFile = _layerFolder + "\\" + layerFileName;
                }

                // Apply layer symbology and move to group layer.
                if (!await SetLayerInMapAsync(layerName, symbologyFile, -1))
                {
                    _searchErrors = true;

                    return false;
                }

                // If labels are to be displayed.
                if (addSelectedLayersOption == AddSelectedLayersOptions.WithLabels && displayLabels)
                {
                    // Translate the label string.
                    if (!string.IsNullOrEmpty(labelClause) && string.IsNullOrEmpty(layerFileName)) // Only if we don't have a layer file.
                    {
                        try
                        {
                            List<string> labelOptions = [.. labelClause.Split('$')];
                            string labelFont = labelOptions[0].Split(':')[1];
                            double labelSize = double.Parse(labelOptions[1].Split(':')[1]);
                            int labelRed = int.Parse(labelOptions[2].Split(':')[1]);
                            int labelGreen = int.Parse(labelOptions[3].Split(':')[1]);
                            int labelBlue = int.Parse(labelOptions[4].Split(':')[1]);
                            string labelOverlap = labelOptions[5].Split(':')[1];
                            bool allowOverlap = labelOverlap.ToLower() switch
                            {
                                "allow" => true,
                                _ => false,
                            };

                            if (await _mapFunctions.LabelLayerAsync(layerName, labelColumn, labelFont, labelSize, "Normal",
                                labelRed, labelGreen, labelBlue, allowOverlap))
                                FileFunctions.WriteLine(_logFile, "Labels added to output " + layerName);
                        }
                        catch
                        {
                            //MessageBox.Show("Error adding labels to '" + layerName + "'");
                            FileFunctions.WriteLine(_logFile, "Error adding labels to '" + layerName + "'");
                            _searchErrors = true;

                            return false;
                        }
                    }
                    else if (!string.IsNullOrEmpty(labelColumn) && string.IsNullOrEmpty(layerFileName))
                    {
                        // Set simple labels.
                        if (await _mapFunctions.LabelLayerAsync(layerName, labelColumn))
                            FileFunctions.WriteLine(_logFile, "Labels added to output " + layerName);
                    }
                }
                else
                {
                    // Turn labels off.
                    await _mapFunctions.SwitchLabelsAsync(layerName, displayLabels);
                }
            }
            else
            {
                // User doesn't want to add the layer to the display.
                // In case it's still there from a previous run.
                await _mapFunctions.RemoveLayerAsync(layerName);
            }

            return true;
        }

        /// <summary>
        /// Trigger the required VB macro to post-process the outputs for the
        /// current layer.
        /// </summary>
        /// <param name="macroName"></param>
        /// <param name="mapTableOutputName"></param>
        /// <param name="mapFormat"></param>
        /// <returns></returns>
        public bool StartProcess(string macroName, string mapTableOutputName, string mapFormat)
        {
            using Process scriptProc = new();

            scriptProc.StartInfo.FileName = @"cscript.exe";
            scriptProc.StartInfo.WorkingDirectory = FileFunctions.GetDirectoryName(macroName); //<---very important
            scriptProc.StartInfo.UseShellExecute = true;
            scriptProc.StartInfo.Arguments = string.Format(@"//B //Nologo {0} {1} {2} {3}", "\"" + macroName + "\"", "\"" + _outputFolder + "\"", "\"" + mapTableOutputName + "." + mapFormat.ToLower() + "\"", "\"" + mapTableOutputName + ".xlsx" + "\"");
            scriptProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden; //prevent console window from popping up

            try
            {
                scriptProc.Start();
                scriptProc.WaitForExit(); // <-- Optional if you want program running until your script exits.

                int exitcode = scriptProc.ExitCode;
                if (exitcode != 0)
                {
                    FileFunctions.WriteLine(_logFile, "Error executing vbscript macro. Exit code : " + exitcode);
                    _searchErrors = true;

                    return false;
                }
            }
            catch
            {
                return false;
            }
            finally
            {
                scriptProc.Close();
            }

            return true;
        }

        #endregion Methods

        #region Debugging Aides

        /// <summary>
        /// Warns the developer if this object does not have
        /// a public property with the specified name. This
        /// method does not exist in a Release build.
        /// </summary>
        [Conditional("DEBUG")]
        [DebuggerStepThrough]
        public void VerifyPropertyName(string propertyName)
        {
            // Verify that the property name matches a real,
            // public, instance property on this object.
            if (TypeDescriptor.GetProperties(this)[propertyName] == null)
            {
                string msg = "Invalid property name: " + propertyName;

                if (ThrowOnInvalidPropertyName)
                    throw new(msg);
                else
                    Debug.Fail(msg);
            }
        }

        /// <summary>
        /// Returns whether an exception is thrown, or if a Debug.Fail() is used
        /// when an invalid property name is passed to the VerifyPropertyName method.
        /// The default value is false, but subclasses used by unit tests might
        /// override this property's getter to return true.
        /// </summary>
        protected virtual bool ThrowOnInvalidPropertyName { get; private set; }

        #endregion Debugging Aides

        #region INotifyPropertyChanged Members

        /// <summary>
        /// Raised when a property on this object has a new value.
        /// </summary>
        public new event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Raises this object's PropertyChanged event.
        /// </summary>
        /// <param name="propertyName">The property that has a new value.</param>
        internal virtual void OnPropertyChanged(string propertyName)
        {
            VerifyPropertyName(propertyName);

            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                PropertyChangedEventArgs e = new(propertyName);
                handler(this, e);
            }
        }

        #endregion INotifyPropertyChanged Members
    }

    /// <summary>
    /// Map layers to search.
    /// </summary>
    public class SQLLayer : INotifyPropertyChanged
    {
        #region Fields

        public string NodeName { get; set; }

        public string NodeGroup { get; set; }

        public string NodeLayer { get; set; }

        public string OutputName { get; set; }

        public string OutputType { get; set; }

        public string Columns { get; set; }

        public string WhereClause { get; set; }

        public string OrderColumns { get; set; }

        public bool LoadWarning { get; set; }

        public string MacroName { get; set; }

        public string MacroParms { get; set; }

        private bool _isSelected;

        public bool IsSelected
        {
            get { return _isSelected; }
            set
            {
                _isSelected = value;

                OnPropertyChanged(nameof(IsSelected));
            }
        }

        #endregion Fields

        #region Creator

        public SQLLayer()
        {
            // constructor takes no arguments.
        }

        public SQLLayer(string nodeName)
        {
            NodeName = nodeName;
        }

        #endregion Creator

        #region INotifyPropertyChanged Members

        /// <summary>
        /// Raised when a property on this object has a new value.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Raises this object's PropertyChanged event.
        /// </summary>
        /// <param name="propertyName">The property that has a new value.</param>
        internal virtual void OnPropertyChanged(string propertyName)
        {
            //VerifyPropertyName(propertyName);

            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                PropertyChangedEventArgs e = new(propertyName);
                handler(this, e);
            }
        }

        #endregion INotifyPropertyChanged Members
    }

    /// <summary>
    /// Map layers to search.
    /// </summary>
    public class MapLayer : INotifyPropertyChanged
    {
        #region Fields

        public string NodeName { get; set; }

        public string NodeGroup { get; set; }

        public string NodeLayer { get; set; }

        public string LayerName { get; set; }

        public string OutputName { get; set; }

        public string OutputType { get; set; }

        public string Columns { get; set; }

        public string WhereClause { get; set; }

        public string OrderColumns { get; set; }

        public bool LoadWarning { get; set; }

        public string MacroName { get; set; }

        public string MacroParms { get; set; }

        private bool _isSelected;

        public bool IsSelected
        {
            get { return _isSelected; }
            set
            {
                _isSelected = value;

                OnPropertyChanged(nameof(IsSelected));
            }
        }

        #endregion Fields

        #region Creator

        public MapLayer()
        {
            // constructor takes no arguments.
        }

        public MapLayer(string nodeName)
        {
            NodeName = nodeName;
        }

        #endregion Creator

        #region INotifyPropertyChanged Members

        /// <summary>
        /// Raised when a property on this object has a new value.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Raises this object's PropertyChanged event.
        /// </summary>
        /// <param name="propertyName">The property that has a new value.</param>
        internal virtual void OnPropertyChanged(string propertyName)
        {
            //VerifyPropertyName(propertyName);

            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                PropertyChangedEventArgs e = new(propertyName);
                handler(this, e);
            }
        }

        #endregion INotifyPropertyChanged Members
    }
}
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
using ArcGIS.Core.Data.UtilityNetwork.Trace;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Internal.Framework.Controls;
using ArcGIS.Desktop.Internal.KnowledgeGraph;
using ArcGIS.Desktop.Internal.Mapping.CommonControls;
using ArcGIS.Desktop.Internal.Mapping.Controls.QueryBuilder.SqlEditor;
using ArcGIS.Desktop.Layouts;
using ArcGIS.Desktop.Mapping;
using DataTools;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using MessageBox = ArcGIS.Desktop.Framework.Dialogs.MessageBox;

namespace DataExtractor.UI
{
    internal class PaneHeader2ViewModel : PanelViewModelBase, INotifyPropertyChanged
    {
        #region Enums

        /// <summary>
        /// An enumeration of the different options for the selection.
        /// </summary>
        public enum SelectionTypeOptions
        {
            SpatialOnly,
            TagsOnly,
            SpatialAndTags
        };

        #endregion Enums

        #region Fields

        private readonly DockpaneMainViewModel _dockPane;

        private string _sdeFileName;

        private bool _extractErrors;

        private string _logFilePath;
        private string _logFile;

        private string _defaultPath;
        private string _partnerFolder;
        private string _gdbName;
        private string _arcGISFolder;
        private string _csvFolder;
        private string _txtFolder;

        private string _defaultSchema;

        // SQL table fields.
        private string _spatialStoredProcedure;
        private string _subsetStoredProcedure;
        private string _clearSpatialStoredProcedure;
        private string _clearSubsetStoredProcedure;
        private string _includeWildcard;
        private string _excludeWildcard;

        // Partner table fields.
        private string _partnerTable;
        private string _partnerColumn;
        private string _shortColumn;
        private string _notesColumn;
        private string _activeColumn;
        private string _formatColumn;
        private string _exportColumn;
        private string _sqlTableColumn;
        private string _sqlFilesColumn;
        private string _mapFilesColumn;
        private string _tagsColumn;
        private string _spatialColumn;

        private List<Partner> _partners;
        private List<SQLLayer> _sqlLayers;
        private List<MapLayer> _mapLayers;

        private List<string> _sqlTableNames;

        private int _defaultSelectType;
        private string _exclusionClause;
        private bool? _defaultApplyExclusionClause;
        private bool? _defaultUseCentroids;
        private bool? _defaultUploadToServer;

        private string _userID;

        private string _dateDD;
        private string _dateMM;
        private string _dateMMM;
        private string _dateMMMM;
        private string _dateYY;
        private double _dateQtr;
        private string _dateQQ;
        private string _dateYYYY;
        private string _dateFFFF;

        private long _pointCount = 0;
        private long _polyCount = 0;
        private long _tableCount = 0;

        private int _extractCnt;
        private int _extractTot;

        private const string _displayName = "DataExtractor";

        private readonly DataExtractorConfig _toolConfig;
        private MapFunctions _mapFunctions;
        private SQLServerFunctions _sqlFunctions;

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
        /// Initialise the extract pane.
        /// </summary>
        private void InitializeComponent()
        {
            // Set the SDE file name.
            _sdeFileName = _toolConfig.SDEFile;

            // Create a new map functions object.
            _mapFunctions = new();

            // Create a new SQL functions object.
            _sqlFunctions = new(_sdeFileName);

            // Get the relevant config file settings.
            _logFilePath = _toolConfig.LogFilePath;

            _spatialStoredProcedure = _toolConfig.SpatialStoredProcedure;
            _subsetStoredProcedure = _toolConfig.SubsetStoredProcedure;
            _clearSpatialStoredProcedure = _toolConfig.ClearSpatialStoredProcedure;
            _clearSubsetStoredProcedure = _toolConfig.ClearSubsetStoredProcedure;

            _defaultPath = _toolConfig.DefaultPath;
            _partnerFolder = _toolConfig.PartnerFolder;
            _gdbName = _toolConfig.GDBName;
            _arcGISFolder = _toolConfig.ArcGISFolder;
            _csvFolder = _toolConfig.CSVFolder;
            _txtFolder = _toolConfig.TXTFolder;

            _defaultSchema = _toolConfig.DatabaseSchema;

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
            _sqlTableColumn = _toolConfig.SQLTableColumn;
            _sqlFilesColumn = _toolConfig.SQLFilesColumn;
            _mapFilesColumn = _toolConfig.MapFilesColumn;
            _tagsColumn = _toolConfig.TagsColumn;
            _spatialColumn = _toolConfig.SpatialColumn;

            _selectionTypeList = _toolConfig.SelectTypeOptions;
            _defaultSelectType = _toolConfig.DefaultSelectType; ;

            _exclusionClause = _toolConfig.ExclusionClause;
            _defaultApplyExclusionClause = _toolConfig.DefaultApplyExclusionClause;
            _defaultUseCentroids = _toolConfig.DefaultUseCentroids;
            _defaultUploadToServer = _toolConfig.DefaultUploadToServer;
        }

        #endregion Creator

        #region Controls Enabled

        /// <summary>
        /// Is the list of partners enabled?
        /// </summary>
        public bool PartnersListEnabled
        {
            get
            {
                return ((_dockPane.ProcessStatus == null)
                    && (_partnersList != null));
            }
        }

        /// <summary>
        /// Is the list of SQL layers enabled?
        /// </summary>
        public bool SQLLayersListEnabled
        {
            get
            {
                return ((_dockPane.ProcessStatus == null)
                    && (_sqlLayersList != null));
            }
        }

        /// <summary>
        /// Is the list of Map layers enabled?
        /// </summary>
        public bool MapLayersListEnabled
        {
            get
            {
                return ((_dockPane.ProcessStatus == null)
                    && (_sqlLayersList != null));
            }
        }

        /// <summary>
        /// Is the list of selection type options enabled?
        /// </summary>
        public bool SelectionTypeListEnabled
        {
            get
            {
                return ((_dockPane.ProcessStatus == null)
                    && (_selectionTypeList != null));
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
                    && (_partnersList != null)
                    && (_partnersList.Where(p => p.IsSelected).Any())
                    && (((_sqlLayersList != null)
                    && (_sqlLayersList.Where(p => p.IsSelected).Any()))
                    || ((_mapLayersList != null)
                    && (_mapLayersList.Where(p => p.IsSelected).Any())))
                    && (_defaultSelectType <= 0 || _selectionType != null));
            }
        }

        #endregion Controls Enabled

        #region Controls Visibility

        /// <summary>
        /// Is the PartnersList expand button visible.
        /// </summary>
        public Visibility PartnersListExpandButtonVisibility
        {
            get
            {
                if ((_partnersList == null) || (_partnersList.Count < 9))
                    return Visibility.Hidden;
                else
                    return Visibility.Visible;
            }
        }

        /// <summary>
        /// Is the SQLLayersList expand button visible.
        /// </summary>
        public Visibility SQLLayersListExpandButtonVisibility
        {
            get
            {
                if ((_sqlLayersList == null) || (_sqlLayersList.Count < 9))
                    return Visibility.Hidden;
                else
                    return Visibility.Visible;
            }
        }

        /// <summary>
        /// Is the MapLayersList expand button visible.
        /// </summary>
        public Visibility MapLayersListExpandButtonVisibility
        {
            get
            {
                if ((_mapLayersList == null) || (_mapLayersList.Count < 9))
                    return Visibility.Hidden;
                else
                    return Visibility.Visible;
            }
        }

        /// <summary>
        /// Is the extract apply exclusion clause check box visible.
        /// </summary>
        public Visibility ApplyExclusionClauseVisibility
        {
            get
            {
                if (!_defaultApplyExclusionClause == null)
                    return Visibility.Collapsed;
                else
                    return Visibility.Visible;
            }
        }

        /// <summary>
        /// Is the use centroids check box visible.
        /// </summary>
        public Visibility UseCentroidsVisibility
        {
            get
            {
                if (!_defaultUseCentroids == null)
                    return Visibility.Collapsed;
                else
                    return Visibility.Visible;
            }
        }

        /// <summary>
        /// Is the upload to server check box visible.
        /// </summary>
        public Visibility UploadToServerVisibility
        {
            get
            {
                if (!_defaultUploadToServer == null)
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
        /// Run the extract.
        /// </summary>
        public async void RunExtractAsync()
        {
            // Reset the cancel flag.
            _dockPane.ExtractCancelled = false;

            // Validate the parameters.
            if (!ValidateParameters())
                return;

            // Clear any messages.
            ClearMessage();

            // Replace any illegal characters in the user name string.
            _userID = StringFunctions.StripIllegals(Environment.UserName, "_", false);

            // User ID should be something at least.
            if (string.IsNullOrEmpty(_userID))
                _userID = "Temp";

            // Set the destination log file path.
            _logFile = _logFilePath + @"\DataExtractor_" + _userID + ".log";

            // Archive the log file (if it exists).
            if (ClearLogFile)
            {
                if (FileFunctions.FileExists(_logFile))
                {
                    // Get the last modified date/time for the current log file.
                    DateTime dateLogFile = File.GetLastWriteTime(_logFile);
                    string dateLastMod = dateLogFile.ToString("yyyy") + dateLogFile.ToString("MM") + dateLogFile.ToString("dd") + "_" +
                        dateLogFile.ToString("HH") + dateLogFile.ToString("mm") + dateLogFile.ToString("ss");

                    // Rename the current log file.
                    string logFileArchive = _logFilePath + @"\DataExtractor_" + _userID + "_" + dateLastMod + ".log";
                    if (!FileFunctions.RenameFile(_logFile, logFileArchive))
                    {
                        MessageBox.Show("Error: Cannot rename log file. Please make sure it is not open in another window.", _displayName, MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                }
            }

            // Create log file path.
            if (!FileFunctions.DirExists(_logFilePath))
            {
                try
                {
                    Directory.CreateDirectory(_logFilePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Cannot create directory " + _logFilePath + ". System error: " + ex.Message);
                    return;
                }
            }

            // If userid is temp.
            if (_userID == "Temp")
                FileFunctions.WriteLine(_logFile, "User ID not found. User ID used will be 'Temp'.");

            // Update the fields and buttons in the form.
            UpdateFormControls();
            _dockPane.RefreshPanel1Buttons();

            // Process the extracts.
            bool success = await ProcessExtractsAsync();

            // Indicate that the extract process has completed (successfully or not).
            string message;
            string image;

            if (success)
            {
                message = "Process complete!";
                image = "Success";
            }
            else if (_extractErrors)
            {
                message = "Process ended with errors!";
                image = "Error";
            }
            else if (_dockPane.ExtractCancelled)
            {
                message = "Process cancelled!";
                image = "Warning";
            }
            else
            {
                message = "Process ended unexpectedly!";
                image = "Error";
            }

            // Finish up now the extract has stopped (successfully or not).
            StopExtract(message, image);

            // Reset the cancel flag.
            _dockPane.ExtractCancelled = false;

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
            // At least one partner must be selected,
            if (!PartnersList.Where(p => p.IsSelected).Any())
            {
                ShowMessage("Please select at least one partner to extract.", MessageType.Warning);
                return false;
            }

            // At least one GIS Map layer or SQL table must be selected,
            if (!MapLayersList.Where(p => p.IsSelected).Any() && !SQLLayersList.Where(s => s.IsSelected).Any())
            {
                ShowMessage("Please select at least one map or SQL table to extract.", MessageType.Warning);
                return false;
            }

            // A selection type must be selected.
            if (string.IsNullOrEmpty(SelectionType))
            {
                ShowMessage("Please select whether the combined sites table should be created.", MessageType.Warning);
                return false;
            }

            ClearMessage();
            return true;
        }

        #endregion Run Command

        #region Properties

        /// <summary>
        /// The list of active partners.
        /// </summary>
        private ObservableCollection<Partner> _partnersList;
        private ObservableCollection<Partner> _activePartnersList;

        /// <summary>
        /// Get the list of active partners.
        /// </summary>
        public ObservableCollection<Partner> PartnersList
        {
            get
            {
                return _partnersList;
            }
            set
            {
                _partnersList = value;
                OnPropertyChanged(nameof(PartnersListExpandButtonVisibility));
            }
        }

        private double? _partnersListHeight = null;

        public double? PartnersListHeight
        {
            get
            {
                if (_activePartnersList == null || _activePartnersList.Count == 0)
                    return Double.NaN;
                else
                    return _partnersListHeight;
            }
        }

        public string PartnersListExpandButtonContent
        {
            get
            {
                if (_partnersListHeight == null)
                    return "-";
                else
                    return "+";
            }
        }

        /// <summary>
        /// Triggered when the selection in the list of Partners changes.
        /// </summary>
        public int PartnersList_SelectedIndex
        {
            set
            {
                // Check if the run button is now enabled/disabled.
                _dockPane.CheckRunButton();
            }
        }

        private List<Partner> _selectedPartners;

        /// <summary>
        /// Get/Set the selected SQL tables.
        /// </summary>
        public List<Partner> SelectedPartners
        {
            get
            {
                return _selectedPartners;
            }
            set
            {
                _selectedPartners = value;
            }
        }

        /// <summary>
        /// The list of SQL tables.
        /// </summary>
        private ObservableCollection<SQLLayer> _sqlLayersList;
        private ObservableCollection<SQLLayer> _sqlXMLLayersList;

        /// <summary>
        /// Get the list of SQL tables.
        /// </summary>
        public ObservableCollection<SQLLayer> SQLLayersList
        {
            get
            {
                return _sqlLayersList;
            }
            set
            {
                _sqlLayersList = value;
                OnPropertyChanged(nameof(SQLLayersListExpandButtonVisibility));
            }
        }

        private double? _sqlLayersListHeight = null;

        public double? SQLLayersListHeight
        {
            get
            {
                if (_sqlLayersList == null || _sqlLayersList.Count == 0)
                    return Double.NaN;
                else
                    return _sqlLayersListHeight;
            }
        }

        public string SQLLayersListExpandButtonContent
        {
            get
            {
                if (_sqlLayersListHeight == null)
                    return "-";
                else
                    return "+";
            }
        }

        /// <summary>
        /// Triggered when the selection in the list of SQL tables changes.
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
        /// Get/Set the selected SQL tables.
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
        /// The list of loaded GIS Map layers.
        /// </summary>
        private ObservableCollection<MapLayer> _mapLayersList;
        private ObservableCollection<MapLayer> _openMapLayersList;
        private List<string> _closedMapLayersList;

        /// <summary>
        /// Get the list of loaded GIS Map layers.
        /// </summary>
        public ObservableCollection<MapLayer> MapLayersList
        {
            get
            {
                return _mapLayersList;
            }
            set
            {
                _mapLayersList = value;
                OnPropertyChanged(nameof(MapLayersListExpandButtonVisibility));
            }
        }

        private double? _mapLayersListHeight = null;

        public double? MapLayersListHeight
        {
            get
            {
                if (_openMapLayersList == null || _openMapLayersList.Count == 0)
                    return Double.NaN;
                else
                    return _mapLayersListHeight;
            }
        }

        public string MapLayersListExpandButtonContent
        {
            get
            {
                if (_mapLayersListHeight == null)
                    return "-";
                else
                    return "+";
            }
        }

        /// <summary>
        /// Triggered when the selection in the list of GIS Map layers changes.
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

        private List<string> _selectionTypeList;

        /// <summary>
        /// Get/Set the options for whether the output layers will be added
        /// to the map.
        /// </summary>
        public List<string> SelectionTypeList
        {
            get
            {
                return _selectionTypeList;
            }
            set
            {
                _selectionTypeList = value;

                // Update the fields and buttons in the form.
                OnPropertyChanged(nameof(SelectionTypeListEnabled));
            }
        }

        private string _selectionType;

        /// <summary>
        /// Get/Set the selection type.
        /// </summary>
        public string SelectionType
        {
            get
            {
                return _selectionType;
            }
            set
            {
                _selectionType = value;

                // Check if the run button is now enabled/disabled.
                _dockPane.CheckRunButton();
            }
        }

        private bool _applyExclusionClause;

        /// <summary>
        /// Get/Set if the exclusion clause should be applied.
        /// </summary>
        public bool ApplyExclusionClause
        {
            get
            {
                return _applyExclusionClause;
            }
            set
            {
                _applyExclusionClause = value;
            }
        }

        private bool _useCentroids;

        /// <summary>
        /// Get/Set if centroids should be used for the spatial selection.
        /// </summary>
        public bool UseCentroids
        {
            get
            {
                return _useCentroids;
            }
            set
            {
                _useCentroids = value;
            }
        }

        private bool _uploadToServer;

        /// <summary>
        /// Get/Set if the partner layer should be uploaded to the server.
        /// </summary>
        public bool UploadToServer
        {
            get
            {
                return _uploadToServer;
            }
            set
            {
                _uploadToServer = value;
            }
        }

        private string _partnerClause;

        public string PartnerClause
        {
            get
            {
                return _partnerClause;
            }
            set
            {
                _partnerClause = value;
            }
        }

        private bool _clearLogFile;

        /// <summary>
        /// Is the log file to be cleared before running the extract?
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
        /// Is the log file to be opened after running the extract?
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
            OnPropertyChanged(nameof(PartnersList));
            OnPropertyChanged(nameof(PartnersListEnabled));
            OnPropertyChanged(nameof(SQLLayersList));
            OnPropertyChanged(nameof(SQLLayersListEnabled));
            OnPropertyChanged(nameof(MapLayersList));
            OnPropertyChanged(nameof(SelectionTypeList));
            OnPropertyChanged(nameof(SelectionTypeListEnabled));
            OnPropertyChanged(nameof(Message));
            OnPropertyChanged(nameof(HasMessage));
        }

        /// <summary>
        /// Set all of the form fields to their default values.
        /// </summary>
        public async Task ResetFormAsync(bool reset)
        {
            // Clear the partner selections first (to avoid selections being retained).
            if (_partnersList != null)
            {
                foreach (Partner layer in _partnersList)
                {
                    layer.IsSelected = false;
                }
            }

            // Clear the SQL table selections first (to avoid selections being retained).
            if (_sqlLayersList != null)
            {
                foreach (SQLLayer layer in _sqlLayersList)
                {
                    layer.IsSelected = false;
                }
            }

            // Clear the map layer selections first (to avoid selections being retained).
            if (_mapLayersList != null)
            {
                foreach (MapLayer layer in _mapLayersList)
                {
                    layer.IsSelected = false;
                }
            }

            // Default selection type.
            if (_toolConfig.DefaultSelectType > 0)
                SelectionType = _toolConfig.SelectTypeOptions[_toolConfig.DefaultSelectType - 1];

            // Log file.
            ClearLogFile = _toolConfig.DefaultClearLogFile;
            OpenLogFile = _toolConfig.DefaultOpenLogFile;

            // Pause map.
            PauseMap = _toolConfig.PauseMap;

            // Reload the list of partners, SQL tables, and open GIS map layers.
            await LoadListsAsync(reset, true);
        }

        /// <summary>
        /// Load the list of partners, SQL tables, and open GIS map layers.
        /// </summary>
        /// <param name="reset"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public async Task LoadListsAsync(bool reset, bool message)
        {
            // If already processing then exit.
            if (_dockPane.ProcessStatus != null)
                return;

            // Expand the lists (ready to be resized later).
            _partnersListHeight = null;
            _sqlLayersListHeight = null;
            _mapLayersListHeight = null;

            _dockPane.FormListsLoading = true;
            if (reset)
                _dockPane.ProgressUpdate("Refreshing lists...");
            else
                _dockPane.ProgressUpdate("Loading lists...");

            // Clear any messages.
            ClearMessage();

            // Update the fields and buttons in the form.
            UpdateFormControls();

            // Get the list of SQL table names from SQL Server (don't wait).
            Task sqlTableNamesTask = GetSQLTableNamesAsync();

            // Load the list of partners (don't wait)
            Task<string> partnersTask = LoadPartnersAsync();

            // Reload the list of SQL layers from the XML profile (don't wait).
            Task<string> sqlLayersTask = LoadSQLLayersAsync();

            // Reload the list of GIS map layers (don't wait).
            Task<string> mapLayersTask = LoadMapLayersAsync();

            // Wait for all of the lists to load.
            await Task.WhenAll(partnersTask, sqlTableNamesTask, sqlLayersTask, mapLayersTask);

            // Set the list of active partners (in partner name order).
            PartnersList = new ObservableCollection<Partner>(_activePartnersList.OrderBy(a => a.PartnerName));

            // Set the list of SQL tables.
            SQLLayersList = _sqlXMLLayersList;

            // Set the list of open layers.
            MapLayersList = _openMapLayersList;

            // Hide progress update.
            _dockPane.ProgressUpdate(null, -1, -1);

            // Indicate the refresh has finished.
            _dockPane.FormListsLoading = false;

            // Update the fields and buttons in the form.
            UpdateFormControls();
            _dockPane.RefreshPanel1Buttons();

            // Force list column widths to reset.
            PartnersListExpandCommandClick(null);
            SQLLayersListExpandCommandClick(null);
            MapLayersListExpandCommandClick(null);

            // Show any message from loading the partner list.
            if (partnersTask.Result != null!)
            {
                ShowMessage(partnersTask.Result, MessageType.Warning);
                if (message)
                    MessageBox.Show(partnersTask.Result, _displayName, MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Show any message from loading the SQL layers list.
            if (sqlLayersTask.Result != null!)
            {
                ShowMessage(sqlLayersTask.Result, MessageType.Warning);
                if (message)
                    MessageBox.Show(sqlLayersTask.Result, _displayName, MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // Show any message from loading the map layers list.
            if (mapLayersTask.Result != null!)
            {
                ShowMessage(mapLayersTask.Result, MessageType.Warning);
                if (message)
                    MessageBox.Show(mapLayersTask.Result, _displayName, MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
        }

        /// <summary>
        /// Load the list of active partners.
        /// </summary>
        /// <returns></returns>
        public async Task<string> LoadPartnersAsync()
        {
            if (_mapFunctions == null || _mapFunctions.MapName == null || MapView.Active.Map.Name != _mapFunctions.MapName)
            {
                // Create a new map functions object.
                _mapFunctions = new();
            }

            // Check if there is an active map.
            bool mapOpen = _mapFunctions.MapName != null;

            // Reset the list of active partners.
            _activePartnersList = [];

            if (!mapOpen)
                return null;

            // Check the partner table is loaded.
            if (_mapFunctions.FindLayer(_partnerTable) == null)
                return "Partner table '" + _partnerTable + "' not found.";

            // Check all of the partner columns are in the partner table.
            List<string> allPartnerColumns = _toolConfig.AllPartnerColumns;

            // Get the list of partner columns that exist in the partner table.
            List<string> existingPartnerColumns = await _mapFunctions.GetExistingFieldsAsync(_partnerTable, allPartnerColumns);

            // Report on the fields that aren't found.
            var missingPartnerColumns = allPartnerColumns.Except(existingPartnerColumns).ToList();
            if (missingPartnerColumns.Count != 0)
            {
                string errMessage = "";
                foreach (string columnName in missingPartnerColumns)
                {
                    errMessage = errMessage + "'" + columnName + "', ";
                }
                return string.Format("The column(s) {0} could not be found in table {1}.", errMessage.Substring(0, errMessage.Length - 2), _partnerTable);
            }

            // Set the default partner where clause
            _partnerClause = _toolConfig.PartnerClause;
            if (String.IsNullOrEmpty(_partnerClause))
                _partnerClause = _toolConfig.ActiveColumn + " = 'Y'";

            // Get the list of active partners from the partner layer.
            _partners = await _mapFunctions.GetActiveParnersAsync(_partnerTable, _partnerClause, _partnerColumn, _shortColumn, _notesColumn,
                _formatColumn, _exportColumn, _sqlTableColumn, _sqlFilesColumn, _mapFilesColumn, _tagsColumn, _activeColumn);

            // Show a message if there are no active partners.
            if (_partners == null || _partners.Count == 0)
                return string.Format("No active partners found in table {0}", _partnerTable);

            // Loop through all of the active partners and add them
            // to the list.
            foreach (Partner partner in _partners)
            {
                // Add the active partners to the list.
                _activePartnersList.Add(partner);
            }

            return null;
        }

        /// <summary>
        /// Load the list of SQL layers.
        /// </summary>
        /// <param name="message"></param>
        /// <returns></returns>
        public async Task<string> LoadSQLLayersAsync()
        {
            // Reset the list of SQL tables.
            _sqlXMLLayersList = [];

            // Load the SQL table variables from the XML profile.
            try
            {
                await Task.Run(() =>
                {
                if (!_toolConfig.GetSQLVariables())
                    return;
                });
            }
            catch (Exception)
            {
                // Only report message if user was prompted for the XML
                // file (i.e. the user interface has already loaded).
                return "Error loading SQL variables from XML file.";
            }

            // Get all of the SQL table details.
            _sqlLayers = _toolConfig.SQLLayers;

            await Task.Run(() =>
            {
                // Loop through all of the layers to check if they are open
                // in the active map.
                foreach (SQLLayer table in _sqlLayers)
                {
                    // Add the open layers to the list.
                    _sqlXMLLayersList.Add(table);
                }
            });

            return null;
        }

        /// <summary>
        /// Load the list of open GIS layers.
        /// </summary>
        /// <param name="message"></param>
        /// <returns></returns>
        public async Task<string> LoadMapLayersAsync()
        {
            // Reset the list of open layers.
            _openMapLayersList = [];

            // Rest the list of closed layers.
            _closedMapLayersList = [];

            // Load the map layer variables from the XML profile.
            try
            {
                await Task.Run(() =>
                {
                    if (!_toolConfig.GetMapVariables())
                        return;
                });
            }
            catch (Exception)
            {
                // Only report message if user was prompted for the XML
                // file (i.e. the user interface has already loaded).
                return "Error loading Map variables from XML file.";
            }

            // Get all of the map layer details.
            _mapLayers = _toolConfig.MapLayers;

            await Task.Run(() =>
            {
                if (_mapFunctions == null || _mapFunctions.MapName == null || MapView.Active.Map.Name != _mapFunctions.MapName)
                {
                    // Create a new map functions object.
                    _mapFunctions = new();
                }

                // Check if there is an active map.
                bool mapOpen = _mapFunctions.MapName != null;

                if (mapOpen)
                {
                    List<MapLayer> allLayers = _mapLayers;

                    // Loop through all of the layers to check if they are open
                    // in the active map.
                    foreach (MapLayer layer in allLayers)
                    {
                        if (_mapFunctions.FindLayer(layer.LayerName) != null)
                        {
                            // Add the open layers to the list.
                            _openMapLayersList.Add(layer);
                        }
                        else
                        {
                            // Only add if the user wants to be warned of this one.
                            if (layer.LoadWarning)
                                _closedMapLayersList.Add(layer.LayerName);
                        }
                    }
                }
            });

            // Show a message if there are no open map layers.
            if (!_openMapLayersList.Any())
                return "No map layers in active map.";

            // Warn the user of any closed map layers.
            int closedLayerCount = _closedMapLayersList.Count;
            if (closedLayerCount > 0)
            {
                string closedLayerWarning = "";
                if (closedLayerCount == 1)
                {
                    closedLayerWarning = "Layer '" + _closedMapLayersList[0] + "' is not loaded.";
                }
                else
                {
                    closedLayerWarning = string.Format("{0} map layers are not loaded.", closedLayerCount.ToString());
                }

                return closedLayerWarning;
            }

            return null;
        }

        /// <summary>
        /// Clear the list of partners, SQL tables, and open GIS map layers.
        /// </summary>
        /// <returns></returns>
        public void ClearFormLists()
        {
            // If not already processing.
            if (_dockPane.ProcessStatus == null)
            {
                // Clear the list of partners.
                PartnersList = [];

                // Clear the list of SQL tables.
                SQLLayersList = [];

                // Clear the list of open GIS map layers.
                MapLayersList = [];

                // Update the fields and buttons in the form.
                UpdateFormControls();
            }
        }

        /// <summary>
        /// Validate and run the extract.
        /// </summary>
        private async Task<bool> ProcessExtractsAsync()
        {
            if (_mapFunctions == null || _mapFunctions.MapName == null)
            {
                // Create a new map functions object.
                _mapFunctions = new();
            }

            // Reset extract errors flag.
            _extractErrors = false;

            // Selected list items.
            _selectedPartners = PartnersList.Where(p => p.IsSelected).ToList();
            _selectedSQLLayers = SQLLayersList.Where(s => s.IsSelected).ToList();
            _selectedMapLayers = MapLayersList.Where(m => m.IsSelected).ToList();

            // What is the selection type?
            SelectionTypeOptions selectionTypeOption = SelectionTypeOptions.SpatialOnly;
            int selectionTypeNum = 0;
            if (_defaultSelectType > 0)
            {
                if (SelectionType.Equals("spatial only", StringComparison.OrdinalIgnoreCase))
                {
                    selectionTypeOption = SelectionTypeOptions.SpatialOnly;
                    selectionTypeNum = 1;
                }
                else if (SelectionType.Equals("survey tags only", StringComparison.OrdinalIgnoreCase))
                {
                    selectionTypeOption = SelectionTypeOptions.TagsOnly;
                    selectionTypeNum = 2;
                }
                else if (SelectionType.Equals("spatial and survey tags", StringComparison.OrdinalIgnoreCase))
                {
                    selectionTypeOption = SelectionTypeOptions.SpatialAndTags;
                    selectionTypeNum = 3;
                }
            }

            // Will the exclusion clause be applied?
            bool applyExclusionClause = ApplyExclusionClause;

            // Will centroids be user for the selection?
            bool useCentroids = UseCentroids;

            // Will the partner table be uploaded to SQL server first?
            bool uploadToServer = UploadToServer;

            // Set the date variables.
            DateTime dateNow = DateTime.Now;
            _dateDD = dateNow.ToString("dd");
            _dateMM = dateNow.ToString("MM");
            _dateMMM = dateNow.ToString("MMM");
            _dateMMMM = dateNow.ToString("MMMM");
            _dateYY = dateNow.ToString("yy");
            _dateQtr = (Math.Ceiling(dateNow.Month / 3.0 + 2) % 4) + 1;
            _dateQQ = _dateQtr.ToString("00");
            _dateYYYY = dateNow.ToString("yyyy");
            _dateFFFF = StringFunctions.FinancialYear(dateNow);

            // Replace any date variables in the log file path.
            string logFilePath = _logFilePath.Replace("%dd%", _dateDD).Replace("%mm%", _dateMM).Replace("%mmm%", _dateMMM).Replace("%mmmm%", _dateMMMM);
            logFilePath = logFilePath.Replace("%yy%", _dateYY).Replace("%qq%", _dateQQ).Replace("%yyyy%", _dateYYYY).Replace("%ffff%", _dateFFFF);

            // Replace any date variables in the default path.
            string defaultPath = _defaultPath.Replace("%dd%", _dateDD).Replace("%mm%", _dateMM).Replace("%mmm%", _dateMMM).Replace("%mmmm%", _dateMMMM);
            defaultPath = defaultPath.Replace("%yy%", _dateYY).Replace("%qq%", _dateQQ).Replace("%yyyy%", _dateYYYY).Replace("%ffff%", _dateFFFF);

            // Replace any date variables in the partner folder.
            string partnerFolder = _partnerFolder.Replace("%dd%", _dateDD).Replace("%mm%", _dateMM).Replace("%mmm%", _dateMMM).Replace("%mmmm%", _dateMMMM);
            partnerFolder = partnerFolder.Replace("%yy%", _dateYY).Replace("%qq%", _dateQQ).Replace("%yyyy%", _dateYYYY).Replace("%ffff%", _dateFFFF);

            // Replace any date variables in the GDB name.
            string gdbName = _gdbName.Replace("%dd%", _dateDD).Replace("%mm%", _dateMM).Replace("%mmm%", _dateMMM).Replace("%mmmm%", _dateMMMM);
            gdbName = gdbName.Replace("%yy%", _dateYY).Replace("%qq%", _dateQQ).Replace("%yyyy%", _dateYYYY).Replace("%ffff%", _dateFFFF);

            // Trim any trailing spaces (directory functions don't deal with them well).
            partnerFolder = partnerFolder.Trim();
            gdbName = gdbName.Trim();
            string arcGISFolder = _arcGISFolder.Trim();
            string csvFolder = _csvFolder.Trim();
            string txtFolder = _txtFolder.Trim();

            // Set a default GDB name if it's empty.
            if (String.IsNullOrEmpty(gdbName))
                gdbName = "Data";

            // Count the number of partners to process.
            int stepsMax = SelectedPartners.Count;
            int stepNum = 0;

            // Count the total number of steps to process.
            _extractCnt = 0;
            _extractTot = SelectedPartners.Count * (SelectedSQLLayers.Count + SelectedMapLayers.Count);

            // Stop if the user cancelled the process.
            if (_dockPane.ExtractCancelled)
                return false;

            // Indicate the extract has started.
            _dockPane.ExtractCancelled = false;
            _dockPane.ExtractRunning = true;

            // Write the first line to the log file.
            FileFunctions.WriteLine(_logFile, "-----------------------------------------------------------------------");
            FileFunctions.WriteLine(_logFile, "Process started!");
            FileFunctions.WriteLine(_logFile, "-----------------------------------------------------------------------");

            // Clear the partner features selection.
            await _mapFunctions.ClearLayerSelectionAsync(_partnerTable);

            // If at least one SQL table is selected).
            if (_selectedSQLLayers.Count > 0)
            {
                if (useCentroids)
                    FileFunctions.WriteLine(_logFile, "Selecting features using centroids ...");
                else
                    FileFunctions.WriteLine(_logFile, "Selecting features using boundaries ...");

                FileFunctions.WriteLine(_logFile, "Performing selection type '" + selectionTypeOption.ToString() + "' ...");

                // Upload the partner table if required.
                if (uploadToServer)
                {
                    _dockPane.ProgressUpdate("Uploading table to server", -1, -1);

                    FileFunctions.WriteLine(_logFile, "Uploading partner table to server ...");

                    if (!await ArcGISFunctions.CopyFeaturesAsync(_partnerTable, _sdeFileName + @"\" + _defaultSchema + "." + _partnerTable, false))
                    {
                        FileFunctions.WriteLine(_logFile, "Error: Uploading partner table.");
                        _extractErrors = true;
                        return false;
                    }

                    FileFunctions.WriteLine(_logFile, "Upload to server complete.");
                }
            }

            // Pause the map redrawing.
            if (PauseMap)
                _mapFunctions.PauseDrawing(true);

            foreach (Partner selectedPartner in SelectedPartners)
            {
                // Stop if the user cancelled the process.
                if (_dockPane.ExtractCancelled)
                    break;

                // Get the partner name and abbreviation.
                string partnerName = selectedPartner.PartnerName;
                string partnerAbbr = selectedPartner.ShortName;

                _dockPane.ProgressUpdate("Processing '" + partnerName + "'...", stepNum, stepsMax);
                stepNum += 1;

                FileFunctions.WriteLine(_logFile, "");
                FileFunctions.WriteLine(_logFile, "----------------------------------------------------------------------");
                FileFunctions.WriteLine(_logFile, "Processing partner '" + partnerName + "' (" + partnerAbbr + ") ...");

                // Loop through the partners, processing each one.
                if (!await ProcessPartnerAsync(selectedPartner, selectionTypeNum, applyExclusionClause, useCentroids, defaultPath, partnerFolder, gdbName, arcGISFolder, csvFolder, txtFolder))
                    _extractErrors = true;

                // Log the completion of this partner.
                FileFunctions.WriteLine(_logFile, "Processing for partner complete.");
                FileFunctions.WriteLine(_logFile, "----------------------------------------------------------------------");
            }

            // Increment the progress value to the last step.
            _dockPane.ProgressUpdate("Cleaning up...", stepNum, 0);

            // Clean up after the extract.
            await CleanUpExtractAsync();

            // If there were errors or the process was cancelled then exit.
            if ((_extractErrors) || (_dockPane.ExtractCancelled))
                return false;

            return true;
        }

        /// <summary>
        /// Indicate that the extract process has stopped (either
        /// successfully or otherwise).
        /// </summary>
        /// <param name="message"></param>
        /// <param name="image"></param>
        private void StopExtract(string message, string image)
        {
            FileFunctions.WriteLine(_logFile, "---------------------------------------------------------------------------");
            FileFunctions.WriteLine(_logFile, message);
            FileFunctions.WriteLine(_logFile, "---------------------------------------------------------------------------");

            // Resume the map redrawing.
            _mapFunctions.PauseDrawing(false);

            // Indicate extract has finished.
            _dockPane.ExtractRunning = false;
            _dockPane.ProgressUpdate(null, -1, -1);

            string imageSource = string.Format("pack://application:,,,/DataExtractor;component/Images/{0}32.png", image);

            // Notify user of completion.
            Notification notification = new()
            {
                Title = _displayName,
                Severity = Notification.SeverityLevel.High,
                Message = message,
                ImageSource = new BitmapImage(new Uri(imageSource)) as ImageSource
            };
            FrameworkApplication.AddNotification(notification);

            // Open the log file (if required).
            if (OpenLogFile || _extractErrors)
                Process.Start("notepad.exe", _logFile);
        }

        /// <summary>
        /// Clean up after the extract has completed (successfully or not).
        /// </summary>
        /// <returns></returns>
        private async Task CleanUpExtractAsync()
        {
            FileFunctions.WriteLine(_logFile, "");

            // Clear the partner features selection.
            await _mapFunctions.ClearLayerSelectionAsync(_partnerTable);
        }

        private async Task<bool> ProcessPartnerAsync(Partner partner, int selectionTypeNum, bool applyExclusionClause, bool useCentroids,
            string defaultPath, string partnerFolder, string gdbName, string arcGISFolder, string csvFolder, string txtFolder)
        {
            // Get the partner details.
            string partnerName = partner.PartnerName;
            string partnerAbbr = partner.ShortName;
            string sqlTable = partner.SQLTable;

            // Select the correct partner polygon.
            string filterClause = _partnerColumn + " = '" + partnerName + "' AND (" + _partnerClause + ")";
            if (!await _mapFunctions.SelectLayerByAttributesAsync(_partnerTable, filterClause, SelectionCombinationMethod.New))
            {
                FileFunctions.WriteLine(_logFile, "Error: Selecting partner boundary.");
                return false;
            }

            // Get the partner table feature layer.
            FeatureLayer partnerLayer = _mapFunctions.FindLayer(_partnerTable);
            if (partnerLayer == null)
                return false;

            // Count the selected features.
            if (partnerLayer.SelectionCount == 0)
            {
                FileFunctions.WriteLine(_logFile, "Error: Partner not found in partner table.");
                return false;
            }
            else if (partnerLayer.SelectionCount > 1)
            {
                FileFunctions.WriteLine(_logFile, "Error: Duplicate entries for partner '" + partnerName + "' in the partner table.");
                return false;
            }

            // Replace the partner shortname in the output names.
            partnerFolder = Regex.Replace(partnerFolder, "%partner%", partnerAbbr, RegexOptions.IgnoreCase);
            gdbName = Regex.Replace(gdbName, "%partner%", partnerAbbr, RegexOptions.IgnoreCase);

            // Set up the output folder.
            string outFolder = defaultPath + @"\" + partnerFolder;
            if (!FileFunctions.DirExists(outFolder))
            {
                FileFunctions.WriteLine(_logFile, "Creating output path '" + outFolder + "'.");
                try
                {
                    Directory.CreateDirectory(outFolder);
                }
                catch (Exception ex)
                {
                    FileFunctions.WriteLine(_logFile, "Error: Cannot create directory '" + outFolder + "'. System error: " + ex.Message);
                    return false;
                }
            }

            //------------------------------------------------------------------
            // Let's start the SQL layers.
            //------------------------------------------------------------------

            Task spatialSelectionTask = null;

            // Set the SQL output table name.
            string sqlOutputTable = sqlTable + "_" + _userID;

            // If at least one SQL table is selected).
            if (_selectedSQLLayers.Count > 0)
            {
                // Check the partner SQL table is found on the server.
                if (string.IsNullOrEmpty(sqlTable))
                {
                    FileFunctions.WriteLine(_logFile, "Skipping SQL outputs - table name is blank.");
                    return true;
                }

                if (!_sqlTableNames.Contains(sqlTable))
                {
                    FileFunctions.WriteLine(_logFile, "Skipping SQL outputs - table '" + sqlTable + "' - not found.");
                    return true;
                }

                // Trigger the spatial selection (but don't wait for it to finish).
                spatialSelectionTask = PerformSpatialSelectionAsync(partner, _defaultSchema, selectionTypeNum, useCentroids);
            }

            //------------------------------------------------------------------
            // Let's do the GIS layers (while the spatial selection runs).
            //------------------------------------------------------------------

            // If at least one map layer is selected).
            if (_selectedMapLayers.Count > 0)
            {
                foreach (MapLayer mapLayer in _selectedMapLayers)
                {
                    _extractCnt += 1;
                    FileFunctions.WriteLine(_logFile, "Starting process " + _extractCnt + " of " + _extractTot + " ...");

                    // Process the required outputs from the GIS layers.
                    if (!await ProcessMapLayerAsync(partner, mapLayer, outFolder, gdbName, arcGISFolder, csvFolder, txtFolder))
                    {
                        // Continue but flag the error.
                        _extractErrors = true;
                    }

                    // Clear the selection in the input layer.
                    await _mapFunctions.ClearLayerSelectionAsync(mapLayer.LayerName);

                    FileFunctions.WriteLine(_logFile, "Completed process " + _extractCnt + " of " + _extractTot + ".");
                    FileFunctions.WriteLine(_logFile, "");
                }
            }

            //------------------------------------------------------------------
            // Let's finish the SQL layers.
            //------------------------------------------------------------------

            // If at least one SQL table is selected).
            if (_selectedSQLLayers.Count > 0)
            {
                // Wait for the spatial selection to complete.
                await spatialSelectionTask;

                FileFunctions.WriteLine(_logFile, "SQL spatial selection complete.");

                // Check if the output feature class exists.
                if (!await _sqlFunctions.FeatureClassExistsAsync(sqlOutputTable))
                {
                    FileFunctions.WriteLine(_logFile, "Procedure returned no records.");
                    return false;
                }

                // Count the number of rows in the output feature count.
                long tableCount = await _sqlFunctions.FeatureClassCountRowsAsync(sqlOutputTable);

                if (tableCount > 0)
                    FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", tableCount) + " records selected.");
                else
                {
                    FileFunctions.WriteLine(_logFile, "Procedure returned no records.");

                    // Clean up the subsets tables in case any are left over.
                    await ClearSubsetTablesAsync(_defaultSchema, sqlTable, _userID);

                    // Clean up the sptial table after processing the partner.
                    await ClearSpatialTableAsync(_defaultSchema, sqlTable, _userID);

                    return false;
                }

                foreach (SQLLayer sqlLayer in _selectedSQLLayers)
                {
                    _extractCnt += 1;
                    FileFunctions.WriteLine(_logFile, "Starting process " + _extractCnt + " of " + _extractTot + " ...");

                    // Process the required outputs from the spatial selection.
                    if (!await ProcessSQLLayerAsync(partner, sqlLayer, applyExclusionClause, outFolder, gdbName, arcGISFolder, csvFolder, txtFolder))
                    {
                        // Continue but flag the error.
                        _extractErrors = true;
                    }

                    FileFunctions.WriteLine(_logFile, "Completed process " + _extractCnt + " of " + _extractTot + ".");
                    FileFunctions.WriteLine(_logFile, "");
                }

                // Clean up the subsets tables in case any are left over.
                await ClearSubsetTablesAsync(_defaultSchema, sqlTable, _userID);

                // Clean up the sptial table after processing the partner.
                await ClearSpatialTableAsync(_defaultSchema, sqlTable, _userID);
            }

            return true;
        }

        private async Task<bool> ProcessSQLLayerAsync(Partner partner, SQLLayer sqlLayer, bool applyExclusionClause, string outFolder,
            string gdbName, string arcGISFolder, string csvFolder, string txtFolder)
        {
            // Get the partner details.
            string partnerAbbr = partner.ShortName;
            string gisFormat = partner.GISFormat?.ToUpper().Trim();
            string exportFormat = partner.ExportFormat?.ToUpper().Trim();
            string sqlTable = partner.SQLTable;
            string sqlFiles = partner.SQLFiles;

            // Get the SQL layer details.
            string nodeName = sqlLayer.NodeName;
            string outputTable = sqlLayer.OutputName;
            string outputType = sqlLayer.OutputType?.ToUpper().Trim();
            string outputColumns = sqlLayer.Columns;
            string whereClause = sqlLayer.WhereClause;
            string orderClause = sqlLayer.OrderColumns;
            string macroName = sqlLayer.MacroName;
            string macroParm = sqlLayer.MacroParms;

            // Check the partner requires something.
            if (((string.IsNullOrEmpty(gisFormat)) && (string.IsNullOrEmpty(exportFormat)) && (string.IsNullOrEmpty(outputType))) || (string.IsNullOrEmpty(sqlFiles)))
            {
                FileFunctions.WriteLine(_logFile, "Skipping output = '" + nodeName + "' - not required.");
                return true;
            }

            // Does the partner want this layer?
            if (!sqlFiles.Contains(nodeName, StringComparison.CurrentCultureIgnoreCase) && !sqlFiles.Equals("all", StringComparison.CurrentCultureIgnoreCase))
            {
                FileFunctions.WriteLine(_logFile, "Skipping output = '" + nodeName + "' - not required.");
                return true;
            }

            // Replace any date variables in the output name.
            outputTable = outputTable.Replace("%dd%", _dateDD).Replace("%mm%", _dateMM).Replace("%mmm%", _dateMMM).Replace("%mmmm%", _dateMMMM);
            outputTable = outputTable.Replace("%yy%", _dateYY).Replace("%qq%", _dateQQ).Replace("%yyyy%", _dateYYYY).Replace("%ffff%", _dateFFFF);

            // Replace the partner shortname in the output name.
            outputTable = Regex.Replace(outputTable, "%partner%", partnerAbbr, RegexOptions.IgnoreCase);

            // Set the output file names.
            string pointFeatureClass = _defaultSchema + "." + sqlTable + "_" + _userID + "_point";
            string polyFeatureClass = _defaultSchema + "." + sqlTable + "_" + _userID + "_poly";
            string flatTable = _defaultSchema + "." + sqlTable + "_" + _userID + "_flat";

            string inPoints = _sdeFileName + @"\" + pointFeatureClass;
            string inPolys = _sdeFileName + @"\" + polyFeatureClass;
            string inFlatTable = _sdeFileName + @"\" + flatTable;

            FileFunctions.WriteLine(_logFile, "Processing output '" + nodeName + "' ...");

            // Is the output spatial and do we need to split for polys/points?
            bool isSpatial = false;
            bool isSplit = false;

            // Check if there is a geometry field in the input
            // table that will be returned.
            string inputTable = _defaultSchema + "." + sqlTable + "_" + _userID;
            string spatialColumn = await IsSQLTableSpatialAsync(inputTable, outputColumns);

            // Set if the output will be spatial and split.
            if (spatialColumn != null)
            {
                isSpatial = true;
                isSplit = true;
            }

            // Add the exclusion clause if required.
            if (applyExclusionClause && !String.IsNullOrEmpty(_exclusionClause))
            {
                if (String.IsNullOrEmpty(whereClause))
                    whereClause = _exclusionClause;
                else
                    whereClause = "(" + whereClause + ") AND (" + _exclusionClause + ")";
            }

            // Write the parameters to the log file.
            FileFunctions.WriteLine(_logFile, string.Format("Column names are '{0}'.", outputColumns));
            if (!string.IsNullOrEmpty(whereClause))
                FileFunctions.WriteLine(_logFile, string.Format("Where clause is '{0}'.", whereClause.Replace("\r\n", " ")));
            else
                FileFunctions.WriteLine(_logFile, "No where clause is specified.");
            if (!string.IsNullOrEmpty(orderClause))
                FileFunctions.WriteLine(_logFile, string.Format("Order by clause is '{0}'.", orderClause.Replace("\r\n", " ")));
            else
                FileFunctions.WriteLine(_logFile, "No order by clause is specified.");
            if (isSplit)
                FileFunctions.WriteLine(_logFile, "Data will be split into points and polygons.");

            // Set the map output format.
            string mapOutputFormat = gisFormat;
            if ((outputType == "GDB") || (outputType == "SHP") || (outputType == "DBF"))
            {
                FileFunctions.WriteLine(_logFile, "Overriding the output type with '" + outputType + "' ...");
                mapOutputFormat = outputType;
            }

            bool checkOutputSize = false;
            if ((mapOutputFormat == "SHP") || (mapOutputFormat == "DBF"))
                checkOutputSize = true;

            // Run the stored procedure to perform the sub-selection.
            // This will split the output into points and polygons if spatial.
            if (!await PerformSubsetSelectionAsync(isSpatial, isSplit, _defaultSchema, sqlTable,
                outputColumns, whereClause, null, orderClause, _userID, checkOutputSize))
            {
                return false;
            }

            // Override the output format if a non-spatial shapefile.
            if ((!isSpatial) && (mapOutputFormat == "SHP"))
                mapOutputFormat = "DBF";

            // Set the output path.
            string outPath = outFolder;
            if (!string.IsNullOrEmpty(arcGISFolder))
                outPath = outFolder + @"\" + arcGISFolder;

            // Set the map output names depending on the output format.
            string outPoints;
            string outPolys;
            string outFlat;
            switch (mapOutputFormat)
            {
                case "GDB":
                    if (isSpatial)
                        mapOutputFormat = "Geodatabase FC";
                    else
                        mapOutputFormat = "Geodatabase Table";

                    outPath = outPath + "\\" + gdbName + ".gdb";
                    outPoints = outPath + @"\" + outputTable + "_Point";
                    outPolys = outPath + @"\" + outputTable + "_Poly";
                    outFlat = outPath + @"\" + outputTable;
                    break;

                case "SHP":
                    mapOutputFormat = "Shapefile";

                    outPoints = outPath + @"\" + outputTable + "_Point.shp";
                    outPolys = outPath + @"\" + outputTable + "_Poly.shp";
                    outFlat = outPath + @"\" + outputTable + ".dbf";
                    break;

                case "DBF":
                    outPoints = outPath + @"\" + outputTable + "_Point.dbf";
                    outPolys = outPath + @"\" + outputTable + "_Poly.dbf";
                    outFlat = outPath + @"\" + outputTable + ".dbf";
                    break;

                default:
                    FileFunctions.WriteLine(_logFile, "Error: Unknown output format '" + mapOutputFormat + "'.");
                    return false;
            }

            // Output the map results if required.
            if (!string.IsNullOrEmpty(mapOutputFormat))
            {
                if (!await CreateSQLOutput(outputTable, mapOutputFormat, outPath, isSpatial, inPoints, inPolys, inFlatTable, outPoints, outPolys, outFlat))
                {
                    // Clean up before returning.
                    await ClearSubsetTablesAsync(_defaultSchema, sqlTable, _userID);

                    return false;
                }
            }

            // Set the export output format.
            string exportOutputFormat = exportFormat;
            if ((outputType == "CSV") || (outputType == "TXT"))
            {
                FileFunctions.WriteLine(_logFile, "Overriding the export type with '" + outputType + "' ...");
                exportOutputFormat = outputType;
            }

            // Reset the output path.
            outPath = outFolder;

            // Set the export output name depending on the output format.
            string expFile;
            switch (exportOutputFormat)
            {
                case "CSV":
                    if (!string.IsNullOrEmpty(csvFolder))
                        outPath = outFolder + @"\" + csvFolder;
                    expFile = outputTable + ".csv";
                    break;

                case "TXT":
                    if (!string.IsNullOrEmpty(txtFolder))
                        outPath = outFolder + @"\" + txtFolder;
                    expFile = outputTable + ".txt";
                    break;

                default:
                    FileFunctions.WriteLine(_logFile, "Error: Unknown output format '" + exportOutputFormat + "'.");

                    // Clean up before returning.
                    await ClearSubsetTablesAsync(_defaultSchema, sqlTable, _userID);

                    return false;
            }

            // Output the text results if required.
            if (!string.IsNullOrEmpty(exportOutputFormat))
            {
                if (!await CreateSQLExport(outputTable, exportOutputFormat, outPath, isSpatial, inPoints, inPolys, inFlatTable, outPath + @"\" + expFile))
                {
                    // Clean up before returning.
                    await ClearSubsetTablesAsync(_defaultSchema, sqlTable, _userID);

                    return false;
                }

                // Trigger the macro if one exists
                if (!string.IsNullOrEmpty(macroName))
                {
                    FileFunctions.WriteLine(_logFile, "Processing the export file ...");

                    if (!StartProcess(macroName, macroParm, outPath, expFile))
                    {
                        FileFunctions.WriteLine(_logFile, "Error: Executing vbscript macro '" + macroName + "'.");

                        // Clean up before returning.
                        await ClearSubsetTablesAsync(_defaultSchema, sqlTable, _userID);

                        return false;
                    }
                }
            }

            // Clean up after processing the layer.
            await ClearSubsetTablesAsync(_defaultSchema, sqlTable, _userID);

            return true;
        }

        private async Task<bool> CreateSQLOutput(string outputTable, string mapOutputFormat, string outPath, bool isSpatial,
            string inPoints, string inPolys, string inFlatTable, string outPoints, string outPolys, string outFlat)
        {
            FileFunctions.WriteLine(_logFile, "Extracting '" + outputTable + "' ...");

            // Create the output path.
            if (!FileFunctions.DirExists(outPath))
            {
                try
                {
                    Directory.CreateDirectory(outPath);
                }
                catch (Exception ex)
                {
                    FileFunctions.WriteLine(_logFile, "Error: Cannot create directory '" + outPath + "'. System error: " + ex.Message);
                    return false;
                }
            }

            // Check the output geodatabase exists.
            if ((mapOutputFormat == "Geodatabase FC") && (!FileFunctions.DirExists(outPath)))
            {
                FileFunctions.WriteLine(_logFile, "Creating output geodatabase ...");

                if (ArcGISFunctions.CreateFileGeodatabase(outPath) == null)
                {
                    FileFunctions.WriteLine(_logFile, "Error: Creating output geodatabase '" + outPath + "'.");
                    return false;
                }

                FileFunctions.WriteLine(_logFile, "Output geodatabase created.");
            }

            // Output the spatial results.
            if (isSpatial)
            {
                if (mapOutputFormat != "DBF")
                {
                    // Export the points.
                    if (_pointCount > 0)
                    {
                        if (!await ArcGISFunctions.CopyFeaturesAsync(inPoints, outPoints, false))
                        {
                            FileFunctions.WriteLine(_logFile, "Error: Outputing '" + inPoints + "' to '" + outPoints + "'.");
                            return false;
                        }

                        // If metadata .xml file exists delete it.
                        string xmlOutFile = outPoints + ".xml";
                        if (FileFunctions.FileExists(xmlOutFile))
                            FileFunctions.DeleteFile(xmlOutFile); // Not checking for success at the moment.
                    }

                    // Export the polygons.
                    if (_polyCount > 0)
                    {
                        if (!await ArcGISFunctions.CopyFeaturesAsync(inPolys, outPolys, false))
                        {
                            FileFunctions.WriteLine(_logFile, "Error: Outputing '" + inPolys + "' to '" + outPolys + "'");
                            return false;
                        }

                        // If metadata .xml file exists delete it.
                        string xmlOutFile = outPolys + ".xml";
                        if (FileFunctions.FileExists(xmlOutFile))
                            FileFunctions.DeleteFile(xmlOutFile); // Not checking for success at the moment.
                    }
                }
                else
                {
                    // Export the points.
                    if (_pointCount > 0)
                    {
                        if (!await ArcGISFunctions.CopyTableAsync(inPoints, outPoints, false))
                        {
                            FileFunctions.WriteLine(_logFile, "Error: Outputing '" + inPoints + "' to '" + outPoints + "'");
                            return false;
                        }

                        // If metadata .xml file exists delete it.
                        string xmlOutFile = outPoints + ".xml";
                        if (FileFunctions.FileExists(xmlOutFile))
                            FileFunctions.DeleteFile(xmlOutFile); // Not checking for success at the moment.
                    }

                    // Export the polygons.
                    if (_polyCount > 0)
                    {
                        if (!await ArcGISFunctions.CopyTableAsync(inPolys, outPolys, false))
                        {
                            FileFunctions.WriteLine(_logFile, "Error: Outputing '" + inPolys + "' to '" + outPolys + "'");
                            return false;
                        }

                        // If metadata .xml file exists delete it.
                        string xmlOutFile = outPolys + ".xml";
                        if (FileFunctions.FileExists(xmlOutFile))
                            FileFunctions.DeleteFile(xmlOutFile); // Not checking for success at the moment.
                    }
                }
            }
            else // Output the non-spatial results.
            {
                if (!await ArcGISFunctions.CopyTableAsync(inFlatTable, outFlat, false))
                {
                    FileFunctions.WriteLine(_logFile, "Error: Outputing '" + inFlatTable + "' to '" + outFlat + "'");
                    return false;
                }

                // If metadata .xml file exists delete it.
                string xmlOutFile = outFlat + ".xml";
                if (FileFunctions.FileExists(xmlOutFile))
                    FileFunctions.DeleteFile(xmlOutFile); // Not checking for success at the moment.
            }

            FileFunctions.WriteLine(_logFile, "Extract complete.");

            return true;
        }

        private async Task<bool> CreateSQLExport(string outputTable, string exportOutputFormat, string outPath,
            bool isSpatial, string inPoints, string inPolys, string inFlatTable, string expFile)
        {
            FileFunctions.WriteLine(_logFile, "Exporting '" + outputTable + "' ...");

            // Create the output path.
            if (!FileFunctions.DirExists(outPath))
            {
                try
                {
                    Directory.CreateDirectory(outPath);
                }
                catch (Exception ex)
                {
                    FileFunctions.WriteLine(_logFile, "Error: Cannot create directory '" + outPath + "'. System error: " + ex.Message);
                    return false;
                }
            }

            // Output the spatial results.
            if (isSpatial)
            {
                // If schema.ini file exists delete it.
                string strIniFile = FileFunctions.GetDirectoryName(expFile) + "\\schema.ini";
                if (FileFunctions.FileExists(strIniFile))
                    FileFunctions.DeleteFile(strIniFile); // Not checking for success at the moment.

                bool blAppend = false;
                // Export the points.
                if (_pointCount > 0)
                {
                    bool result;
                    if (exportOutputFormat == "TXT")
                        result = await _sqlFunctions.CopyToTabAsync(inPoints, expFile, true, false);
                    else
                        result = await _sqlFunctions.CopyToCSVAsync(inPoints, expFile, true, false);

                    if (!result)
                    {
                        FileFunctions.WriteLine(_logFile, "Error: Exporting '" + inPoints + "' to '" + expFile + "'");
                        return false;
                    }

                    blAppend = true;
                }

                // Also export the polygons - append if necessary
                if (_polyCount > 0)
                {
                    bool result;
                    if (exportOutputFormat == "TXT")
                        result = await _sqlFunctions.CopyToTabAsync(inPolys, expFile, true, blAppend);
                    else
                        result = await _sqlFunctions.CopyToCSVAsync(inPolys, expFile, true, blAppend);

                    if (!result)
                    {
                        FileFunctions.WriteLine(_logFile, "Error: Exporting '" + inPolys + "' to '" + expFile + "'");
                        return false;
                    }
                }

                // If metadata .xml file exists delete it.
                string strXmlFile = expFile + ".xml";
                if (FileFunctions.FileExists(strXmlFile))
                    FileFunctions.DeleteFile(strXmlFile); // Not checking for success at the moment.
            }
            else // Output the non-spatial results.
            {
                bool result;
                if (exportOutputFormat == "TXT")
                    result = await _sqlFunctions.CopyToTabAsync(inFlatTable, expFile, false, false);
                else
                    result = await _sqlFunctions.CopyToCSVAsync(inFlatTable, expFile, false, false);

                if (!result)
                {
                    FileFunctions.WriteLine(_logFile, "Error: Exporting '" + inFlatTable + "' to '" + expFile + "'");
                    return false;
                }

                // If metadata .xml file exists delete it.
                string strXmlFile = expFile + ".xml";
                if (FileFunctions.FileExists(strXmlFile))
                    FileFunctions.DeleteFile(strXmlFile); // Not checking for success at the moment.
            }

            FileFunctions.WriteLine(_logFile, "Export complete.");

            return true;
        }
        private async Task<bool> ProcessMapLayerAsync(Partner partner, MapLayer mapLayer, string outFolder,
            string gdbName, string arcGISFolder, string csvFolder, string txtFolder)
        {
            // Get the partner details.
            string partnerAbbr = partner.ShortName;
            string gisFormat = partner.GISFormat?.ToUpper().Trim();
            string exportFormat = partner.ExportFormat?.ToUpper().Trim();
            string mapFiles = partner.MapFiles;

            // Get the SQL layer details.
            string nodeName = mapLayer.NodeName;
            string layerName = mapLayer.LayerName;
            string outputTable = mapLayer.OutputName;
            string outputType = mapLayer.OutputType?.ToUpper().Trim();
            string outputColumns = mapLayer.Columns;
            string whereClause = mapLayer.WhereClause;
            string orderClause = mapLayer.OrderColumns;
            string macroName = mapLayer.MacroName;
            string macroParm = mapLayer.MacroParms;

            // Check the partner requires something.
            if (((string.IsNullOrEmpty(gisFormat)) && (string.IsNullOrEmpty(exportFormat)) && (string.IsNullOrEmpty(outputType))) || (string.IsNullOrEmpty(mapFiles)))
            {
                FileFunctions.WriteLine(_logFile, "Skipping output = '" + nodeName + "' - not required.");
                return true;
            }

            // Does the partner want this layer?
            if (!mapFiles.Contains(nodeName, StringComparison.CurrentCultureIgnoreCase) && !mapFiles.Equals("all", StringComparison.CurrentCultureIgnoreCase))
            {
                FileFunctions.WriteLine(_logFile, "Skipping output = '" + nodeName + "' - not required.");
                return true;
            }

            // Replace any date variables in the output name.
            outputTable = outputTable.Replace("%dd%", _dateDD).Replace("%mm%", _dateMM).Replace("%mmm%", _dateMMM).Replace("%mmmm%", _dateMMMM);
            outputTable = outputTable.Replace("%yy%", _dateYY).Replace("%qq%", _dateQQ).Replace("%yyyy%", _dateYYYY).Replace("%ffff%", _dateFFFF);

            // Replace the partner shortname in the output name.
            outputTable = Regex.Replace(outputTable, "%partner%", partnerAbbr, RegexOptions.IgnoreCase);

            // Build a list of all of the columns required.
            List<string> mapFields = [];
            List<string> rawFields = [.. outputColumns.Split(',')];
            foreach (string mapField in rawFields)
            {
                mapFields.Add(mapField.Trim());
            }

            FileFunctions.WriteLine(_logFile, "Processing output '" + nodeName + "' ...");

            FileFunctions.WriteLine(_logFile, string.Format("Column names are '{0}'.", outputColumns));
            if (!string.IsNullOrEmpty(whereClause))
                FileFunctions.WriteLine(_logFile, string.Format("Where clause is '{0}'.", whereClause.Replace("\r\n", " ")));
            else
                FileFunctions.WriteLine(_logFile, "No where clause is specified.");
            if (!string.IsNullOrEmpty(orderClause))
                FileFunctions.WriteLine(_logFile, string.Format("Order by clause is '{0}'.", orderClause.Replace("\r\n", " ")));
            else
                FileFunctions.WriteLine(_logFile, "No order by clause is specified.");

            // Set the map output format.
            string mapOutputFormat = gisFormat;
            if ((outputType == "GDB") || (outputType == "SHP") || (outputType == "DBF"))
            {
                FileFunctions.WriteLine(_logFile, "Overriding the output type with '" + outputType + "' ...");
                mapOutputFormat = outputType;
            }

            FileFunctions.WriteLine(_logFile, "Executing spatial selection ...");

            // Firstly do the spatial selection.
            if (!await MapFunctions.SelectLayerByLocationAsync(layerName, _partnerTable))
            {
                FileFunctions.WriteLine(_logFile, "Error: Creating selection using spatial query.");
                return false;
            }

            // Find the map layer by name.
            FeatureLayer inputLayer = _mapFunctions.FindLayer(layerName);

            if (inputLayer == null)
            {
                FileFunctions.WriteLine(_logFile, "Error: Finding map layer '" + layerName + "'.");
                return false;
            }

            // Refine the selection by attributes (if required).
            if (inputLayer.SelectionCount > 0 && !string.IsNullOrEmpty(whereClause))
            {
                FileFunctions.WriteLine(_logFile, "Refining selection with criteria " + whereClause + " ...");

                if (!await _mapFunctions.SelectLayerByAttributesAsync(layerName, whereClause, SelectionCombinationMethod.And))
                {
                    FileFunctions.WriteLine(_logFile, "Error: Creating subset selection using attribute query.");
                    return false;
                }
            }

            // Count the selected features.
            int featureCount = inputLayer.SelectionCount;

            // If there is no selection then exit.
            if (featureCount == 0)
            {
                FileFunctions.WriteLine(_logFile, "No features selected in '" + layerName + "'.");
                return true;
            }

            FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", featureCount) + " features selected.");

            // Override the output format if an export is required but no output.
            if ((string.IsNullOrEmpty(gisFormat) && !string.IsNullOrEmpty(exportFormat)))
                mapOutputFormat = "SHP";

            // Set the output path.
            string outPath = outFolder;
            if (!string.IsNullOrEmpty(arcGISFolder))
                outPath = outFolder + @"\" + arcGISFolder;

            // Set the map output names depending on the output format.
            string outFile;
            switch (mapOutputFormat)
            {
                case "GDB":
                    mapOutputFormat = "Geodatabase FC";
                    outFile = outputTable;
                    break;

                case "SHP":
                    mapOutputFormat = "Shapefile";
                    outFile = outputTable + ".shp";
                    break;

                case "DBF":
                    outFile = outputTable + ".dbf";
                    break;

                default:
                    FileFunctions.WriteLine(_logFile, "Error: Unknown output format '" + mapOutputFormat + "'.");
                    return false;
            }

            bool deleteOutput = false;
            if ((gisFormat != outputType) && !string.IsNullOrEmpty(exportFormat))
                deleteOutput = true;

            // Output the map results if required.
            if (!string.IsNullOrEmpty(mapOutputFormat))
            {
                if (!await CreateMapOutput(layerName, outputTable, mapOutputFormat, outPath, gdbName, outFile))
                    return false;

                // Drop non-required fields.
                if (!await _mapFunctions.KeepSelectedFieldsAsync(outputTable, mapFields))
                {
                    FileFunctions.WriteLine(_logFile, "Error: Removing unwanted fields.");
                    {
                        // Remove and delete the feature layer from the map.
                        await ClearMapTablesAsync(outputTable, outPath + @"\" + outFile, deleteOutput);

                        return false;
                    }
                }
            }

            // Set the export output format.
            string exportOutputFormat = exportFormat;
            if ((outputType == "CSV") || (outputType == "TXT"))
            {
                FileFunctions.WriteLine(_logFile, "Overriding the export type with '" + outputType + "' ...");
                exportOutputFormat = outputType;
            }

            // Set the export path.
            string expPath = outFolder;

            // Set the default export file name.
            string expFile;

            // Set the export output name depending on the output format.
            switch (exportOutputFormat)
            {
                case "CSV":
                    if (!string.IsNullOrEmpty(csvFolder))
                        expPath = outFolder + @"\" + csvFolder;
                    expFile = outputTable + ".csv";
                    break;

                case "TXT":
                    if (!string.IsNullOrEmpty(txtFolder))
                        expPath = outFolder + @"\" + txtFolder;
                    expFile = outputTable + ".txt";
                    break;

                default:
                    FileFunctions.WriteLine(_logFile, "Error: Unknown output format '" + exportOutputFormat + "'.");
                    return false;
            }

            // Export the text results if required.
            if (!string.IsNullOrEmpty(exportOutputFormat))
            {
                if (!await CreateMapExport(outputTable, exportOutputFormat, expPath, expFile, outputColumns))
                {
                    // Remove and delete the feature layer from the map.
                    await ClearMapTablesAsync(outputTable, outPath + @"\" + outFile, deleteOutput);

                    return false;
                }

                // Trigger the macro if one exists.
                if (!string.IsNullOrEmpty(macroName))
                {
                    FileFunctions.WriteLine(_logFile, "Processing the export file ...");

                    if (!StartProcess(macroName, macroParm, expPath, expFile))
                    {
                        FileFunctions.WriteLine(_logFile, "Error: Executing vbscript macro '" + macroName + "'.");
                        {
                            // Remove and delete the feature layer from the map.
                            await ClearMapTablesAsync(outputTable, outPath + @"\" + outFile, deleteOutput);

                            return false;
                        }
                    }
                }
            }

            // If the output is in a geodatabase.
            if (mapOutputFormat == "Geodatabase FC")
            {
                // Update the output path to include the geodatabase name.
                outPath = outPath + "\\" + gdbName + ".gdb";
            }

            // Remove and delete the feature layer from the map.
            await ClearMapTablesAsync(outputTable, outPath + @"\" + outFile, deleteOutput);

            return true;
        }

        private async Task<bool> CreateMapOutput(string layerName, string outputTable, string mapOutputFormat,
            string outPath, string gdbName, string outFile)
        {
            FileFunctions.WriteLine(_logFile, "Extracting '" + outputTable + "' ...");

            // Create the output path.
            if (!FileFunctions.DirExists(outPath))
            {
                try
                {
                    Directory.CreateDirectory(outPath);
                }
                catch (Exception ex)
                {
                    FileFunctions.WriteLine(_logFile, "Error: Cannot create directory '" + outPath + "'. System error: " + ex.Message);
                    return false;
                }
            }

            // If the output is in a geodatabase.
            if (mapOutputFormat == "Geodatabase FC")
            {
                // Update the output path to include the geodatabase name.
                outPath = outPath + "\\" + gdbName + ".gdb";

                // Check the output geodatabase exists.
                if (!FileFunctions.DirExists(outPath))
                {
                    FileFunctions.WriteLine(_logFile, "Creating output geodatabase ...");

                    if (ArcGISFunctions.CreateFileGeodatabase(outPath) == null)
                    {
                        FileFunctions.WriteLine(_logFile, "Error: Creating output geodatabase '" + outPath + "'.");
                        return false;
                    }

                    FileFunctions.WriteLine(_logFile, "Output geodatabase created.");
                }
            }

            // Output the features.
            if (!await ArcGISFunctions.CopyFeaturesAsync(layerName, outPath + @"\" + outFile, true))
            {
                FileFunctions.WriteLine(_logFile, "Error: Outputing '" + layerName + "' to '" + outPath + @"\" + outFile + "'.");
                return false;
            }

            // If metadata .xml file exists delete it.
            string xmlOutFile = outPath + @"\" + outFile + ".xml";
            if (FileFunctions.FileExists(xmlOutFile))
                FileFunctions.DeleteFile(xmlOutFile); // Not checking for success at the moment.

            FileFunctions.WriteLine(_logFile, "Extract complete.");

            return true;
        }

        private async Task<bool> CreateMapExport(string outputTable, string exportOutputFormat, string outPath,
            string expFile, string outputColumns)
        {
            FileFunctions.WriteLine(_logFile, "Exporting '" + outputTable + "' ...");

            // Create the output path.
            if (!FileFunctions.DirExists(outPath))
            {
                try
                {
                    Directory.CreateDirectory(outPath);
                }
                catch (Exception ex)
                {
                    FileFunctions.WriteLine(_logFile, "Error: Cannot create directory '" + outPath + "'. System error: " + ex.Message);
                    return false;
                }
            }

            // If schema.ini file exists delete it.
            string strIniFile = FileFunctions.GetDirectoryName(outPath) + "\\schema.ini";
            if (FileFunctions.FileExists(strIniFile))
                FileFunctions.DeleteFile(strIniFile); // Not checking for success at the moment.

            // Export the features.
            int result;
            if (exportOutputFormat == "TXT")
                result = await _mapFunctions.CopyFCToTextFileAsync(outputTable, outPath + @"\" + expFile, outputColumns, null, "\t", false, false);
            else
                result = await _mapFunctions.CopyFCToTextFileAsync(outputTable, outPath + @"\" + expFile, outputColumns, null, ",", false, true);

            if (result <= 0)
            {
                FileFunctions.WriteLine(_logFile, "Error exporting '" + outputTable + "' to '" + outPath + @"\" + expFile + "'.");
                return false;
            }

            // If metadata .xml file exists delete it.
            string strXmlFile = outPath + @"\" + expFile + ".xml";
            if (FileFunctions.FileExists(strXmlFile))
                FileFunctions.DeleteFile(strXmlFile); // Not checking for success at the moment.

            FileFunctions.WriteLine(_logFile, "Export complete.");
            return true;
        }

        /// <summary>
        /// Perform the spatial selection via a stored procedure.
        /// </summary>
        /// <param name="partner"></param>
        /// <param name="schema"></param>
        /// <param name="selectionTypeNum"></param>
        /// <param name="useCentroids"></param>
        /// <returns></returns>
        internal Task PerformSpatialSelectionAsync(Partner partner, string schema, int selectionTypeNum, bool useCentroids)
        {
            // Get the partner details.
            string partnerAbbr = partner.ShortName;
            string sqlTable = partner.SQLTable;

            // Get the name of the stored procedure to execute selection in SQL Server.
            string storedProcedureName = _spatialStoredProcedure;

            // Set up the SQL command.
            StringBuilder sqlCmd = new();

            // Build the SQL command to execute the stored procedure.
            sqlCmd = sqlCmd.Append(string.Format("EXECUTE {0}", storedProcedureName));
            sqlCmd.Append(string.Format(" '{0}'", schema));
            sqlCmd.Append(string.Format(", '{0}'", _partnerTable));
            sqlCmd.Append(string.Format(", '{0}'", _shortColumn));
            sqlCmd.Append(string.Format(", '{0}'", partnerAbbr));
            sqlCmd.Append(string.Format(", '{0}'", _tagsColumn));
            sqlCmd.Append(string.Format(", '{0}'", _spatialColumn));
            sqlCmd.Append(string.Format(", {0}", selectionTypeNum));
            sqlCmd.Append(string.Format(", '{0}'", sqlTable));
            sqlCmd.Append(string.Format(", '{0}'", _userID));
            sqlCmd.Append(string.Format(", {0}", useCentroids ? "1" : "0"));

            FileFunctions.WriteLine(_logFile, "Executing SQL spatial selection from '" + sqlTable + "' ...");

            // Execute the stored procedure.
            return _sqlFunctions.ExecuteSQLOnGeodatabase(sqlCmd.ToString());
        }

        /// <summary>
        /// Perform the subset selection by running the query via a
        /// stored procedure.
        /// </summary>
        /// <param name="isSpatial"></param>
        /// <param name="isSplit"></param>
        /// <param name="schema"></param>
        /// <param name="tableName"></param>
        /// <param name="columnNames"></param>
        /// <param name="whereClause"></param>
        /// <param name="groupClause"></param>
        /// <param name="orderClause"></param>
        /// <param name="userID"></param>
        /// <returns></returns>
        internal async Task<bool> PerformSubsetSelectionAsync(bool isSpatial, bool isSplit, string schema, string tableName,
                                  string columnNames, string whereClause, string groupClause, string orderClause, string userID,
                                  bool checkOutputSize)
        {
            bool success;

            // Get the name of the stored procedure to execute selection in SQL Server.
            string storedProcedureName = _subsetStoredProcedure;

            // Set up the SQL command.
            StringBuilder sqlCmd = new();

            // Double-up single quotes so they parse in SQL command correctly.
            if (whereClause != null && whereClause.Contains('\''))
                whereClause = whereClause.Replace("'", "''");

            // Build the SQL command to execute the stored procedure.
            sqlCmd = sqlCmd.Append(string.Format("EXECUTE {0}", storedProcedureName));
            sqlCmd.Append(string.Format(" '{0}'", schema));
            sqlCmd.Append(string.Format(", '{0}'", tableName + "_" + userID));
            sqlCmd.Append(string.Format(", '{0}'", columnNames));
            sqlCmd.Append(string.Format(", '{0}'", whereClause));
            sqlCmd.Append(string.Format(", '{0}'", groupClause));
            sqlCmd.Append(string.Format(", '{0}'", orderClause));
            sqlCmd.Append(string.Format(", {0}", isSplit ? "1" : "0"));

            // Reset the counters.
            _pointCount = 0;
            _polyCount = 0;
            _tableCount = 0;

            long maxCount;
            long rowLength = 0;

            // Set the SQL output file names.
            string pointFeatureClass = schema + "." + tableName + "_" + userID + "_point";
            string polyFeatureClass = schema + "." + tableName + "_" + userID + "_poly";
            string flatTable = schema + "." + tableName + "_" + userID + "_flat";

            try
            {
                FileFunctions.WriteLine(_logFile, "Performing selection ...");

                // Execute the stored procedure.
                await _sqlFunctions.ExecuteSQLOnGeodatabase(sqlCmd.ToString());

                // If the result is isSpatial it should be split into points and polys.
                if (isSpatial)
                {
                    // Check if the point or polygon feature class exists.
                    success = await _sqlFunctions.FeatureClassExistsAsync(pointFeatureClass);
                    if (!success)
                        success = await _sqlFunctions.FeatureClassExistsAsync(polyFeatureClass);
                }
                else
                {
                    // Check if the table exists.
                    success = await _sqlFunctions.TableExistsAsync(flatTable);
                }

                // If the result table(s) exist.
                if (success)
                {
                    if (isSpatial)
                    {
                        // Count the number of rows in the point feature count.
                        _pointCount = await _sqlFunctions.FeatureClassCountRowsAsync(pointFeatureClass);

                        // Save the maximum row count.
                        maxCount = _pointCount;

                        // Count the number of rows in the poly feature count.
                        _polyCount = await _sqlFunctions.FeatureClassCountRowsAsync(polyFeatureClass);

                        // Update the maximum row count.
                        if (_pointCount > maxCount)
                            maxCount = _pointCount;

                        if (maxCount == 0)
                        {
                            FileFunctions.WriteLine(_logFile, "No records returned.");
                            return true;
                        }

                        FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", _pointCount) + " points to extract, " + string.Format("{0:n0}", _polyCount) + " polygons to extract.");

                        // Calculate the total row length for the table.
                        if (_pointCount > 0)
                            rowLength = await _sqlFunctions.TableRowLength(pointFeatureClass);
                        else if (_polyCount > 0)
                            rowLength += await _sqlFunctions.TableRowLength(polyFeatureClass);
                    }
                    else
                    {
                        // Count the number of rows in the table.
                        _tableCount = await _sqlFunctions.TableCountRowsAsync(flatTable);

                        // Save the maximum row count.
                        maxCount = _tableCount;
                        if (maxCount == 0)
                        {
                            FileFunctions.WriteLine(_logFile, "No records returned.");
                            return true;
                        }

                        FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", _tableCount) + " records to extract");
                        FileFunctions.WriteLine(_logFile, "Data is not spatial");

                        // Calculate the total row length for the table.
                        if (_tableCount > 0)
                            rowLength = await _sqlFunctions.TableRowLength(flatTable);
                    }
                }
                else
                {
                    FileFunctions.WriteLine(_logFile, "No records to extract.");
                    return true;
                }
            }
            catch (Exception ex)
            {
                FileFunctions.WriteLine(_logFile, "Error: Executing the stored procedure: " + ex.Message);
                return false;
            }

            // Return if no need to check output row/file size.
            if (!checkOutputSize)
                return true;

            // Check if the maximum record length will be exceeded
            if (rowLength > 4000)
            {
                FileFunctions.WriteLine(_logFile, "Error: Record length exceeds maximum of 4,000 bytes.");
                return false;
            }

            // Display the maximum data size.
            long maxDataSize;
            maxDataSize = ((rowLength * maxCount) / 1024) + 1;
            string strDataSizeKb = string.Format("{0:n0}", maxDataSize);
            string strDataSizeMb = string.Format("{0:n2}", (double)maxDataSize / 1024);
            string strDataSizeGb = string.Format("{0:n2}", (double)maxDataSize / (1024 * 1024));

            if (maxDataSize > (1024 * 1024))
            {
                FileFunctions.WriteLine(_logFile, string.Format("Maximum data size = {0} Kb ({1} Gb).", strDataSizeKb, strDataSizeGb));
            }
            else
            {
                if (maxDataSize > 1024)
                    FileFunctions.WriteLine(_logFile, string.Format("Maximum data size = {0} Kb ({1} Mb).", strDataSizeKb, strDataSizeMb));
                else
                    FileFunctions.WriteLine(_logFile, string.Format("Maximum data size = {0} Kb.", strDataSizeKb));
            }

            // Check if the maximum data size will be exceeded.
            if (maxDataSize > (2 * 1024 * 1024))
            {
                FileFunctions.WriteLine(_logFile, "Error: Maximum data size exceeds 2 Gb.");
                return false;
            }

            return true;
        }

        /// <summary>
        /// Clear the temporary SQL spatial table by running a stored procedure.
        /// </summary>
        /// <param name="schema"></param>
        /// <param name="tableName"></param>
        /// <param name="userID"></param>
        /// <returns></returns>
        internal async Task<bool> ClearSpatialTableAsync(string schema, string tableName, string userID)
        {
            // Set up the SQL command.
            StringBuilder sqlCmd = new();

            // Get the name of the stored procedure to clear the
            // spatial selection in SQL Server.
            string clearSpatialSPName = _clearSpatialStoredProcedure;

            // Build the SQL command to execute the stored procedure.
            sqlCmd = sqlCmd.Append(string.Format("EXECUTE {0}", clearSpatialSPName));
            sqlCmd.Append(string.Format(" '{0}'", schema));
            sqlCmd.Append(string.Format(", '{0}'", tableName));
            sqlCmd.Append(string.Format(", '{0}'", userID));

            try
            {
                //FileFunctions.WriteLine(_logFile, "Deleting spatial temporary table");

                // Execute the stored procedure.
                await _sqlFunctions.ExecuteSQLOnGeodatabase(sqlCmd.ToString());
            }
            catch (Exception ex)
            {
                FileFunctions.WriteLine(_logFile, "Error: Deleting the spatial temporary table: " + ex.Message);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Clear the temporary SQL subset tables by running a stored procedure.
        /// </summary>
        /// <param name="schema"></param>
        /// <param name="tableName"></param>
        /// <param name="userID"></param>
        /// <returns></returns>
        internal async Task<bool> ClearSubsetTablesAsync(string schema, string tableName, string userID)
        {
            // Set up the SQL command.
            StringBuilder sqlCmd = new();

            // Get the name of the stored procedure to clear the
            // subset selections in SQL Server.
            string clearSubsetSPName = _clearSubsetStoredProcedure;

            // Build the SQL command to execute the stored procedure.
            sqlCmd = sqlCmd.Append(string.Format("EXECUTE {0}", clearSubsetSPName));
            sqlCmd.Append(string.Format(" '{0}'", schema));
            sqlCmd.Append(string.Format(", '{0}'", tableName));
            sqlCmd.Append(string.Format(", '{0}'", userID));

            try
            {
                //FileFunctions.WriteLine(_logFile, "Deleting subset temporary tables");

                // Execute the stored procedure.
                await _sqlFunctions.ExecuteSQLOnGeodatabase(sqlCmd.ToString());
            }
            catch (Exception ex)
            {
                FileFunctions.WriteLine(_logFile, "Error: Deleting the subset temporary tables: " + ex.Message);
                return false;
            }

            return true;
        }

        internal async Task<bool> ClearMapTablesAsync(string outputTable, string outFile, bool deleteOutput)
        {
            // Remove the feature layer from the map.
            if (!await _mapFunctions.RemoveLayerAsync(outputTable))
            {
                FileFunctions.WriteLine(_logFile, "Error: Removing layer '" + outputTable + "' from map.");
                return false;
            }

            // If the output layer is not required (i.e. only exported as
            // text file).
            if (deleteOutput)
            {
                // Delete the feature class.
                if (!await ArcGISFunctions.DeleteFeatureClassAsync(outFile))
                {
                    FileFunctions.WriteLine(_logFile, "Error: Deleting output file '" + outFile + "'.");
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

        /// <summary>
        /// Trigger the required VB macro to post-process the outputs for the
        /// current layer.
        /// </summary>
        /// <param name="macroName"></param>
        /// <param name="macroParm"></param>
        /// <param name="outPath"></param>
        /// <param name="outFile"></param>
        /// <returns></returns>
        public bool StartProcess(string macroName, string macroParm, string outPath, string outFile)
        {
            using Process scriptProc = new();

            // Set the process parameters.
            scriptProc.StartInfo.FileName = @"wscript.exe";
            scriptProc.StartInfo.WorkingDirectory = FileFunctions.GetDirectoryName(macroName);
            scriptProc.StartInfo.UseShellExecute = true;
            scriptProc.StartInfo.Arguments = string.Format("//B //Nologo \"{0}\" \"{1}\" \"{2}\" \"{3}\" \"{4}\"", macroName, macroParm, outPath, outFile, _logFile);
            scriptProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

            try
            {
                scriptProc.Start();
                scriptProc.WaitForExit(); // <-- Optional if you want program running until your script exits.

                int exitcode = scriptProc.ExitCode;
                if (exitcode != 0)
                {
                    FileFunctions.WriteLine(_logFile, "Error: Executing vbscript macro. Exit code : " + exitcode);
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

        #region PartnersListExpand Command

        private ICommand _partnersListExpandCommand;

        /// <summary>
        /// Create PartnersList Expand button command.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand PartnersListExpandCommand
        {
            get
            {
                if (_partnersListExpandCommand == null)
                {
                    Action<object> expandPartnerListAction = new(PartnersListExpandCommandClick);
                    _partnersListExpandCommand = new RelayCommand(expandPartnerListAction, param => true);
                }
                return _partnersListExpandCommand;
            }
        }

        /// <summary>
        /// Handles event when PartnersListExpand button is pressed.
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void PartnersListExpandCommandClick(object param)
        {
            if (_partnersListHeight == null)
                _partnersListHeight = 179;
            else
                _partnersListHeight = null;

            OnPropertyChanged(nameof(PartnersListHeight));
            OnPropertyChanged(nameof(PartnersListExpandButtonContent));
        }

        #endregion PartnersListExpand Command

        #region SQLLayersListExpand Command

        private ICommand _sqlLayersListExpandCommand;

        /// <summary>
        /// Create SQLLayersList Expand button command.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand SQLLayersListExpandCommand
        {
            get
            {
                if (_sqlLayersListExpandCommand == null)
                {
                    Action<object> expandSQLLayersListAction = new(SQLLayersListExpandCommandClick);
                    _sqlLayersListExpandCommand = new RelayCommand(expandSQLLayersListAction, param => true);
                }
                return _sqlLayersListExpandCommand;
            }
        }

        /// <summary>
        /// Handles event when SQLLayersListExpand button is pressed.
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void SQLLayersListExpandCommandClick(object param)
        {
            if (_sqlLayersListHeight == null)
                _sqlLayersListHeight = 162;
            else
                _sqlLayersListHeight = null;

            OnPropertyChanged(nameof(SQLLayersListHeight));
            OnPropertyChanged(nameof(SQLLayersListExpandButtonContent));
        }

        #endregion SQLLayersListExpand Command

        #region MapLayersListExpand Command

        private ICommand _mapLayersListExpandCommand;

        /// <summary>
        /// Create MapLayersList Expand button command.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand MapLayersListExpandCommand
        {
            get
            {
                if (_mapLayersListExpandCommand == null)
                {
                    Action<object> expandMapLayersListAction = new(MapLayersListExpandCommandClick);
                    _mapLayersListExpandCommand = new RelayCommand(expandMapLayersListAction, param => true);
                }
                return _mapLayersListExpandCommand;
            }
        }

        /// <summary>
        /// Handles event when MapLayersListExpand button is pressed.
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void MapLayersListExpandCommandClick(object param)
        {
            if (_mapLayersListHeight == null)
                _mapLayersListHeight = 162;
            else
                _mapLayersListHeight = null;

            OnPropertyChanged(nameof(MapLayersListHeight));
            OnPropertyChanged(nameof(MapLayersListExpandButtonContent));
        }

        #endregion PartnersListExpand Command

        #region SQL

        /// <summary>
        /// Get a list of the SQL table names from the SQL Server.
        /// </summary>
        /// <returns></returns>
        public async Task GetSQLTableNamesAsync()
        {
            // Get the full list of feature classes and tables from SQL Server.
            await _sqlFunctions.GetTableNamesAsync();

            // Get the list of tables returned from SQL Server.
            List<string> tabList = _sqlFunctions.TableNames;

            // If no tables were found.
            if (_sqlFunctions.TableNames.Count == 0)
            {
                // Clear the tables list.
                _sqlTableNames = [];
                return;
            }

            // Get the include and exclude wildcard settings.
            string includeWC = _includeWildcard;
            string excludeWC = _excludeWildcard;

            // Filter the SQL table names and add them to a list.
            List<string> filteredTableList = FilterTableNames(tabList, _defaultSchema, includeWC, excludeWC, false);

            // Set the tables list in sort order.
            _sqlTableNames = new(filteredTableList.OrderBy(t => t));
        }

        /// <summary>
        /// Filter the list of the table names base on the include and exclude wildcard criteria.
        /// </summary>
        /// <param name="inputNames"></param>
        /// <param name="schema"></param>
        /// <param name="includeWildcard"></param>
        /// <param name="excludeWildcard"></param>
        /// <param name="includeFullName"></param>
        /// <returns></returns>
        internal static List<string> FilterTableNames(List<string> inputNames, string schema, string includeWildcard, string excludeWildcard,
                              bool includeFullName = false)
        {
            // Define the wildcards as case insensitive
            Wildcard theInclude = new(includeWildcard, schema, RegexOptions.IgnoreCase);
            Wildcard theExclude = new(excludeWildcard, schema, RegexOptions.IgnoreCase);

            List<string> theStringList = [];

            foreach (string inName in inputNames)
            {
                string tableName = inName;
                // Does the name conform to the includeWildcard?
                if (theInclude.IsMatch(tableName))
                {
                    if (!theExclude.IsMatch(tableName))
                    {
                        if (includeFullName)
                        {
                            theStringList.Add(tableName);
                        }
                        else
                        {
                            tableName = tableName.Split('.')[1];
                            theStringList.Add(tableName);
                        }
                    }
                }
            }

            return theStringList;
        }

        /// <summary>
        /// Check if the table contains a spatial column in the columns text.
        /// </summary>
        /// <param name="tableName"></param>
        /// <param name="columnsText"></param>
        /// <returns></returns>
        internal async Task<string> IsSQLTableSpatialAsync(string tableName, string columnsText)
        {
            string[] geometryFields = ["SP_GEOMETRY", "Shape"]; // Expand as required.

            // Get the list of field names in the selected table.
            List<string> fieldsList = await _sqlFunctions.GetFieldNamesListAsync(tableName);

            // Loop through the geometry fields looking for a match.
            foreach (string geomField in geometryFields)
            {
                // If the columns text contains the geometry field.
                if (columnsText.Contains(geomField, StringComparison.OrdinalIgnoreCase))
                {
                    return geomField;
                }
                // If "*" is used check for the existence of the geometry field in the table.
                else if (columnsText.Equals("*", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (string fieldName in fieldsList)
                    {
                        // If the column text contains the geometry field.
                        if (fieldName.Equals(geomField, StringComparison.OrdinalIgnoreCase))
                            return geomField;
                    }
                }
            }

            // No geometry column found.
            return null;
        }

        #endregion SQL

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

    #region Partner Class

    /// <summary>
    /// Partner to extract.
    /// </summary>
    public class Partner : INotifyPropertyChanged
    {
        #region Fields

        public string PartnerName { get; set; }

        public string ShortName { get; set; }

        public string GISFormat { get; set; }

        public string ExportFormat { get; set; }

        public string SQLTable { get; set; }

        public string SQLFiles { get; set; }

        public string MapFiles { get; set; }

        public string Tags { get; set; }

        public string Notes { get; set; }

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

        public Partner()
        {
            // constructor takes no arguments.
        }

        public Partner(string partnerName)
        {
            PartnerName = partnerName;
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

    #endregion Partner Class

    #region SQLLayer Class

    /// <summary>
    /// Map layers to extract.
    /// </summary>
    public class SQLLayer : INotifyPropertyChanged
    {
        #region Fields

        public string NodeName { get; set; }

        public string NodeGroup { get; set; }

        public string NodeTable { get; set; }

        public string OutputName { get; set; }

        public string OutputType { get; set; }

        public string Columns { get; set; }

        public string WhereClause { get; set; }

        public string OrderColumns { get; set; }

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

    #endregion SQLLayer Class

    #region MapLayer Class

    /// <summary>
    /// Map layers to extract.
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

    #endregion MapLayer Class
}
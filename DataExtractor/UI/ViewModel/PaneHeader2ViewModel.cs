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

using ArcGIS.Core.Data.UtilityNetwork.Trace;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Internal.Framework.Controls;
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
using System.Diagnostics;
using System.IO;
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

        private List<string> _selectTypeOptions = [];

        private List<Partner> _partners;
        private List<SQLLayer> _sqlLayers;
        private List<MapLayer> _mapLayers;

        private List<string> _sqlTableNames;

        private int _defaultSelectType;
        private string _exclusionClause;
        private bool? _defaultApplyExclusionClause;
        private bool? _defaultUseCentroids;
        private bool? _defaultUploadToServer;
        private bool _defaultClearLogFile;
        private bool _defaultOpenLogFile;

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

            _selectTypeOptions = _toolConfig.SelectTypeOptions;
            _defaultSelectType = _toolConfig.DefaultSelectType; ;

            _exclusionClause = _toolConfig.ExclusionClause;
            _defaultApplyExclusionClause = _toolConfig.DefaultApplyExclusionClause;
            _defaultUseCentroids = _toolConfig.DefaultUseCentroids;
            _defaultUploadToServer = _toolConfig.DefaultUploadToServer;
            _defaultClearLogFile = _toolConfig.DefaultClearLogFile;
            _defaultOpenLogFile = _toolConfig.DefaultOpenLogFile;

            // Get all of the SQL table details.
            _sqlLayers = _toolConfig.SQLLayers;

            // Get all of the map layer details.
            _mapLayers = _toolConfig.MapLayers;
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

                    // Rename the current log file
                    string logFileArchive = _logFilePath + @"\DataExtractor_" + _userID + "_" + dateLastMod + ".log";
                    if (!FileFunctions.RenameFile(_logFile, logFileArchive))
                    {
                        MessageBox.Show("Cannot rename log file. Please make sure it is not open in another window", _displayName, MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                }
            }

            // Create log file path.
            if (!FileFunctions.DirExists(_logFilePath))
            {
                //FileFunctions.WriteLine(_logFile, "Creating output path '" + outFolder + "'.");
                try
                {
                    Directory.CreateDirectory(_logFilePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Cannot create directory " + _logFilePath + ". System error: " + ex.Message);
                    return;
                }
            }

            // If userid is temp.
            if (_userID == "Temp")
                FileFunctions.WriteLine(_logFile, "User ID not found. User ID used will be 'Temp'");

            // Update the fields and buttons in the form.
            UpdateFormControls();
            _dockPane.RefreshPanel1Buttons();

            // Process the extracts.
            bool success = await ProcessExtractsAsync();

            // Indicate that the extract process has completed (successfully or not).
            string message;
            string image;
            //TODO - set message
            if (success)
            {
                message = "Extracts complete!";
                image = "Success";
            }
            else if (_extractErrors)
            {
                message = "Extracts ended with errors!";
                image = "Error";
            }
            else if (_dockPane.ExtractCancelled)
            {
                message = "Extracts cancelled!";
                image = "Warning";
            }
            else
            {
                message = "Extracts ended unexpectedly!";
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

        public string SQLNodeGroupWidth
        {
            get { return _toolConfig.SQLNodeGroupWidth; }
        }

        public string MapNodeGroupWidth
        {
            get { return _toolConfig.MapNodeGroupWidth; }
        }

        /// <summary>
        /// The list of active partners.
        /// </summary>
        private ObservableCollection<Partner> _partnersList;

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

        private int? _partnersListHeight = 179;

        public int? PartnersListHeight
        {
            get
            {
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

        private int? _sqlLayersListHeight = 162;

        public int? SQLLayersListHeight
        {
            get
            {
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

        private int? _mapLayersListHeight = 162;

        public int? MapLayersListHeight
        {
            get
            {
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
            OnPropertyChanged(nameof(MapLayersListEnabled));
            OnPropertyChanged(nameof(SelectionTypeList));
            OnPropertyChanged(nameof(SelectionTypeListEnabled));
            //OnPropertyChanged(nameof(ApplyExclusionClause));
            //OnPropertyChanged(nameof(ApplyExclusionClauseVisibility));
            //OnPropertyChanged(nameof(UseCentroids));
            //OnPropertyChanged(nameof(UseCentroidsVisibility));
            //OnPropertyChanged(nameof(UploadToServer));
            //OnPropertyChanged(nameof(UploadToServerVisibility));
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

            // Selection type options.
            SelectionTypeList = _toolConfig.SelectTypeOptions;
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
            // If not already processing.
            if (_dockPane.ProcessStatus == null)
            {
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
                Task sqlTableNamesTask = GetSQLTableNamesAsync(false);

                // Load the list of partners (don't wait)
                Task partnersTask = LoadPartnersAsync(message);

                // Reload the list of GIS map layers (don't wait).
                Task mapLayersTask = LoadMapLayersAsync(message);

                // Reload the list of SQL layers from the XML profile.
                LoadSQLLayers();

                // Wait for all of the lists to load.
                await Task.WhenAll(partnersTask, sqlTableNamesTask, mapLayersTask);

                // Hide progress update.
                _dockPane.ProgressUpdate(null, -1, -1);

                // Indicate the refresh has finished.
                _dockPane.FormListsLoading = false;

                // Update the fields and buttons in the form.
                UpdateFormControls();
                _dockPane.RefreshPanel1Buttons();
            }
        }

        /// <summary>
        /// Load the list of active partners.
        /// </summary>
        /// <returns></returns>
        public async Task LoadPartnersAsync(bool message)
        {
            if (_mapFunctions == null || _mapFunctions.MapName == null || MapView.Active.Map.Name != _mapFunctions.MapName)
            {
                // Create a new map functions object.
                _mapFunctions = new();
            }

            // Check if there is an active map.
            bool mapOpen = _mapFunctions.MapName != null;

            // Reset the list of active partners.
            ObservableCollection<Partner> partnerList = [];

            if (mapOpen)
            {
                // Check the partner table is loaded.
                if (_mapFunctions.FindLayer(_partnerTable) == null)
                {
                    //TODO
                    if (message)
                        MessageBox.Show("Partner table '" + _partnerTable + "' not found.", _displayName, MessageBoxButton.OK, MessageBoxImage.Error);

                    return;
                }

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
                    errMessage = string.Format("The column(s) {0} could not be found in table {1}.", errMessage.Substring(0, errMessage.Length - 2), _partnerTable);

                    //TODO
                    if (message)
                        MessageBox.Show(errMessage, _displayName, MessageBoxButton.OK, MessageBoxImage.Error);

                    return;
                }

                // Set the default partner where clause
                _partnerClause = _toolConfig.PartnerClause;
                if (String.IsNullOrEmpty(_partnerClause))
                    _partnerClause = _toolConfig.ActiveColumn + " = 'Y'";

                // Get the list of active partners from the partner layer.
                _partners = await _mapFunctions.GetActiveParnersAsync(_partnerTable, _partnerClause, _partnerColumn, _shortColumn, _notesColumn,
                    _formatColumn, _exportColumn, _sqlTableColumn, _sqlFilesColumn, _mapFilesColumn, _tagsColumn, _activeColumn);

                if (_partners == null || _partners.Count == 0)
                {
                    //TODO
                    if (message)
                        MessageBox.Show(string.Format("No active partners found in table {0}", _partnerTable), _displayName, MessageBoxButton.OK, MessageBoxImage.Error);

                    return;
                }

                // Loop through all of the active partners and add them
                // to the list.
                foreach (Partner partner in _partners)
                {
                    // Add the active partners to the list.
                    partnerList.Add(partner);
                }
            }

            // Reset the list of active partners.
            PartnersList = partnerList;
        }

        /// <summary>
        /// Load the list of SQL layers.
        /// </summary>
        /// <returns></returns>
        public void LoadSQLLayers()
        {
            // Reset the list of SQL tables.
            ObservableCollection<SQLLayer> sqlLayers = [];

            // Loop through all of the layers to check if they are open
            // in the active map.
            foreach (SQLLayer table in _sqlLayers)
            {
                // Add the open layers to the list.
                sqlLayers.Add(table);
            }

            // Reset the list of SQL tables.
            SQLLayersList = sqlLayers;
        }

        /// <summary>
        /// Load the list of open GIS layers.
        /// </summary>
        /// <param name="message"></param>
        /// <returns></returns>
        public async Task LoadMapLayersAsync(bool message)
        {
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

                // Set the list of open layers.
                MapLayersList = openLayersList;
            });

            // Show a message if there are no open layers.
            if (!_mapLayersList.Any())
            {
                ShowMessage("No map layers in active map.", MessageType.Warning);
                return;
            }

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
                    MessageBox.Show(closedLayerWarning, _displayName, MessageBoxButton.OK, MessageBoxImage.Warning);
            }
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

            //// Trim any trailing spaces (directory functions don't deal with them well).
            partnerFolder = partnerFolder.Trim();
            string arcGISFolder = _arcGISFolder.Trim();
            string csvFolder = _csvFolder.Trim();
            string txtFolder = _txtFolder.Trim();

            // Count the number of partners to process.
            int stepsMax = SelectedPartners.Count;
            int stepNum = 0;

            //TODO = needed?
            // Count the total number of steps to process.
            //_extractTot = SelectedPartners.Count * (SelectedSQLLayers.Count + SelectedMapLayers.Count);

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

                    //FileFunctions.WriteLine(_logFile, "");
                    FileFunctions.WriteLine(_logFile, "Uploading partner table to server ...");

                    if (!await ArcGISFunctions.CopyFeaturesAsync(_partnerTable, _sdeFileName + @"\" + _defaultSchema + "." + _partnerTable, false))
                    {
                        //MessageBox.Show("Error uploading partner table - process terminated.");
                        FileFunctions.WriteLine(_logFile, "Error uploading partner table - process terminated");
                        _extractErrors = true;
                        return false;
                    }

                    FileFunctions.WriteLine(_logFile, "Upload to server complete");
                }
            }

            // Pause the map redrawing.
            if (PauseMap)
                _mapFunctions.PauseDrawing(true);

            bool success;

            //TODO - needed?
            int layerNum = 0;
            int layerCount = SelectedPartners.Count;

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

                //TODO - needed?
                layerNum += 1;

                FileFunctions.WriteLine(_logFile, "");
                FileFunctions.WriteLine(_logFile, "----------------------------------------------------------------------");
                FileFunctions.WriteLine(_logFile, "Processing partner '" + partnerName + "' (" + partnerAbbr + ") ...");

                // Loop through the partners, processing each one.
                success = await ProcessPartnerAsync(selectedPartner, selectionTypeNum, applyExclusionClause, useCentroids, defaultPath, partnerFolder, arcGISFolder, csvFolder, txtFolder);

                //TODO - needed?
                // Keep track of any errors.
                if (!success)
                    _extractErrors = true;

                // Log the completion of this partner.
                FileFunctions.WriteLine(_logFile, "Processing for partner complete.");
                FileFunctions.WriteLine(_logFile, "----------------------------------------------------------------------");
            }

            FileFunctions.WriteLine(_logFile, "");
            FileFunctions.WriteLine(_logFile, "----------------------------------------------------------------------");
            FileFunctions.WriteLine(_logFile, "Process completed!");
            FileFunctions.WriteLine(_logFile, "----------------------------------------------------------------------");

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

            //TODO
            // Delete the temporary output tables in the SQL database.

            //TODO
            // Delete the final temporary spatial table.
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

        private async Task<bool> ProcessPartnerAsync(Partner partner, int selectionTypeNum, bool applyExclusionClause, bool useCentroids,
            string defaultPath, string partnerFolder, string arcGISFolder, string csvFolder, string txtFolder)
        {
            // Get the partner details.
            string partnerName = partner.PartnerName;
            string partnerAbbr = partner.ShortName;
            string sqlTable = partner.SQLTable;

            // Select the correct partner polygon.
            string filterClause = _partnerColumn + " = '" + partnerName + "' AND " + _partnerClause;
            if (!await _mapFunctions.SelectLayerByAttributesAsync(_partnerTable, filterClause, SelectionCombinationMethod.New))
            {
                //TODO - set error message?
                FileFunctions.WriteLine(_logFile, "Error selecting partner boundary");
                return false;
            }

            // Get the partner table feature layer.
            FeatureLayer partnerLayer = _mapFunctions.FindLayer(_partnerTable);
            if (partnerLayer == null)
                return false;

            // Count the selected features.
            if (partnerLayer.SelectionCount == 0)
            {
                FileFunctions.WriteLine(_logFile, "Partner not found in partner table");
                return false;
            }
            else if (partnerLayer.SelectionCount > 1)
            {
                //TODO - message box needed?
                MessageBox.Show("There are duplicate entries for partner " + partnerName + " in the partner table.");
                FileFunctions.WriteLine(_logFile, "There are duplicate entries for partner " + partnerName + " in the partner table");
            }

            // Replace the partner shortname in the output name
            partnerFolder = Regex.Replace(partnerFolder, "%partner%", partnerAbbr, RegexOptions.IgnoreCase);

            // Set up the output folder.
            string outFolder = defaultPath + @"\" + partnerFolder;
            if (!FileFunctions.DirExists(outFolder))
            {
                //FileFunctions.WriteLine(_logFile, "Creating output path '" + outFolder + "'.");
                try
                {
                    Directory.CreateDirectory(outFolder);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Cannot create directory " + outFolder + ". System error: " + ex.Message);
                    FileFunctions.WriteLine(_logFile, "Cannot create directory " + outFolder + ". System error: " + ex.Message);

                    return false;
                }
            }

            //------------------------------------------------------------------
            // Let's start the SQL layers.
            //------------------------------------------------------------------

            Task<bool> spatialSelectionTask = null;

            // If at least one SQL table is selected).
            if (_selectedSQLLayers.Count > 0)
            {
                // Check the partner SQL table is found on the server.
                if (string.IsNullOrEmpty(sqlTable))
                {
                    FileFunctions.WriteLine(_logFile, "Skipping SQL outputs - table name is blank");
                    return true;
                }

                if (!_sqlTableNames.Contains(sqlTable))
                {
                    FileFunctions.WriteLine(_logFile, "Skipping SQL outputs - table '" + sqlTable + "' - not found");
                    return true;
                }

                // Set the SQL output table name.
                string sqlOutputTable = sqlTable + "_" + _userID;

                // Trigger the spatial selection (but don't wait for it to finish).
                spatialSelectionTask = PerformSpatialSelectionAsync(partner, _defaultSchema, selectionTypeNum, sqlOutputTable, useCentroids);
            }

            //------------------------------------------------------------------
            // Let's do the GIS layers (while the spatial selection runs).
            //------------------------------------------------------------------

            // If at least one map layer is selected).
            if (_selectedMapLayers.Count > 0)
            {
                foreach (MapLayer mapLayer in _selectedMapLayers)
                {
                    //TODO - needed?
                    //_extractCnt = extractCnt + 1;
                    //FileFunctions.WriteLine(_logFile, "Starting process " + intExtractCnt + " of " + intExtractTot + " ...");

                    // Process the required outputs from the GIS layers.
                    if (!await ProcessMapLayerAsync(partner, mapLayer, outFolder, partnerFolder, arcGISFolder, csvFolder, txtFolder))
                    {
                        //TODO - just continue but flag the error?
                        _extractErrors = true;
                    }

                    // Clear the selection in the input layer.
                    await _mapFunctions.ClearLayerSelectionAsync(mapLayer.LayerName);

                    //TODO - needed?
                    //FileFunctions.WriteLine(_logFile, "Completed process " + intExtractCnt + " of " + intExtractTot + ".");
                    //FileFunctions.WriteLine(_logFile, "");

                    //intLayerIndex++;
                }
            }

            //------------------------------------------------------------------
            // Let's finish the SQL layers.
            //------------------------------------------------------------------

            // If at least one SQL table is selected).
            if (_selectedSQLLayers.Count > 0)
            {
                // Wait for the spatial selection to complete.
                if (!await spatialSelectionTask)
                {
                    //TODO - error performing spatial selection.
                    return false;
                }

                foreach (SQLLayer sqlLayer in _selectedSQLLayers)
                {
                    //TODO - needed?
                    //_extractCnt = extractCnt + 1;
                    //FileFunctions.WriteLine(_logFile, "Starting process " + intExtractCnt + " of " + intExtractTot + " ...");

                    // Process the required outputs from the spatial selection.
                    if (!await ProcessSQLLayerAsync(partner, sqlLayer, applyExclusionClause, outFolder, partnerFolder, arcGISFolder, csvFolder, txtFolder))
                    {
                        //TODO - just continue but flag the error?
                        _extractErrors = true;
                    }

                    //TODO - needed?
                    //FileFunctions.WriteLine(_logFile, "Completed process " + intExtractCnt + " of " + intExtractTot + ".");
                    //FileFunctions.WriteLine(_logFile, "");

                    //intLayerIndex++;
                }

                // Clean up after processing the partner.
                await ClearSpatialTableAsync(_defaultSchema, sqlTable, _userID);
            }




            return true;
        }

        private async Task<bool> ProcessSQLLayerAsync(Partner partner, SQLLayer sqlLayer, bool applyExclusionClause, string outFolder, string partnerFolder,
            string arcGISFolder, string csvFolder, string txtFolder)
        {
            // Get the partner details.
            string partnerAbbr = partner.ShortName;
            string gisFormat = partner.GISFormat;
            string exportFormat = partner.ExportFormat;
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
                FileFunctions.WriteLine(_logFile, "Skipping output = '" + nodeName + "' - not required");
                return true;
            }

            // Does the partner want this layer?
            if (!sqlFiles.Contains(nodeName, StringComparison.CurrentCultureIgnoreCase) && !sqlFiles.Equals("all", StringComparison.CurrentCultureIgnoreCase))
            {
                FileFunctions.WriteLine(_logFile, "Skipping output = '" + nodeName + "' - not required");
                return true;
            }

            //TODO - needed?
            //_dockPane.ProgressUpdate("Processing SQL table " + nodeName + " ...", -1, 0);

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

            // Build a list of all of the columns required.
            //TODO - loop needed?
            List<string> mapFields = [];
            List<string> rawFields = [.. outputColumns.Split(',')];
            foreach (string mapField in rawFields)
            {
                mapFields.Add(mapField.Trim());
            }

            FileFunctions.WriteLine(_logFile, "");
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
            FileFunctions.WriteLine(_logFile, string.Format("Source table is '{0}'", sqlTable));
            FileFunctions.WriteLine(_logFile, string.Format("Column names are '{0}'", outputColumns));
            if (!string.IsNullOrEmpty(whereClause))
                FileFunctions.WriteLine(_logFile, string.Format("Where clause is '{0}'", whereClause));
            else
                FileFunctions.WriteLine(_logFile, "No where clause is specified");
            if (!string.IsNullOrEmpty(orderClause))
                FileFunctions.WriteLine(_logFile, string.Format("Order by clause is '{0}'", orderClause));
            else
                FileFunctions.WriteLine(_logFile, "No order by clause is specified");
            if (isSplit)
                FileFunctions.WriteLine(_logFile, "Data will be split into points and polygons");

            // Set the map output format.
            string mapOutputFormat = gisFormat.ToUpper();
            if ((outputType == "GDB") || (outputType == "SHP") || (outputType == "DBF"))
            {
                FileFunctions.WriteLine(_logFile, "Overriding the output type with '" + outputType + "' ...");
                mapOutputFormat = outputType;
            }

            // Run the stored procedure to perform the sub-selection.
            // This splits the output into points and polygons as required.
            if (!await PerformSubsetSelectionAsync(isSpatial, isSplit, _defaultSchema, sqlTable,
                outputColumns, whereClause, null, orderClause, _userID, mapOutputFormat))
            {
                _extractErrors = true;
                return false;
            }

            // Override the output format if a non-spatial shapefile.
            if ((!isSpatial) && (mapOutputFormat == "SHP"))
                mapOutputFormat = "DBF";

            // Set the output path.
            string outPath = outFolder;
            if (!string.IsNullOrEmpty(arcGISFolder))
                outPath = outFolder + @"\" + arcGISFolder;

            // Create the output path.
            if (!FileFunctions.DirExists(outPath))
            {
                //FileFunctions.WriteLine(_logFile, "Creating output path '" + outPath + "'.");
                try
                {
                    Directory.CreateDirectory(outPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Cannot create directory " + outPath + ". System error: " + ex.Message);
                    FileFunctions.WriteLine(_logFile, "Cannot create directory " + outPath + ". System error: " + ex.Message);

                    return false;
                }
            }

            // Set the output file name prefixes.
            string outPoints = outputTable;
            string outPolys = outputTable;
            string outFlat = outputTable;

            // Set the map output names depending on the output format.
            switch (mapOutputFormat)
            {
                case "GDB":
                    if (isSpatial)
                        mapOutputFormat = "Geodatabase FC";
                    else
                        mapOutputFormat = "Geodatabase Table";

                    outPath = outPath + "\\" + partnerFolder + ".gdb";
                    outPoints = outPath + @"\" + outputTable + "_Point";
                    outPolys = outPath + @"\" + outputTable + "_Poly";
                    break;

                case "SHP":
                    mapOutputFormat = "Shapefile";

                    outPoints = outPath + @"\" + outputTable + "_Point.shp";
                    outPolys = outPath + @"\" + outputTable + "_Poly.shp";
                    break;

                case "DBF":
                    outPoints = outPath + @"\" + outputTable + "_Point.dbf";
                    outPolys = outPath + @"\" + outputTable + "_Poly.dbf";
                    outFlat = outPath + @"\" + outputTable + ".dbf";
                    break;

                default:
                    mapOutputFormat = "";
                    break;
            }

            // Output the map results if required.
            if (!string.IsNullOrEmpty(mapOutputFormat))
            {
                //TODO - needed?
                //_dockPane.ProgressUpdate("Writing output for " + nodeName + " to GIS format ...", -1, 0);

                FileFunctions.WriteLine(_logFile, "Extracting '" + outputTable + "' ...");

                // Check the output geodatabase exists.
                if ((mapOutputFormat == "Geodatabase FC") && (!FileFunctions.DirExists(outPath)))
                {
                    FileFunctions.WriteLine(_logFile, "Creating output geodatabase ...");

                    ArcGISFunctions.CreateFileGeodatabase(outPath);

                    FileFunctions.WriteLine(_logFile, "Output geodatabase created");
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
                                FileFunctions.WriteLine(_logFile, "Error outputing " + inPoints + " to " + outPoints);
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
                                FileFunctions.WriteLine(_logFile, "Error outputing " + inPolys + " to " + outPolys);
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
                                FileFunctions.WriteLine(_logFile, "Error outputing " + inPoints + " to " + outPoints);
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
                                FileFunctions.WriteLine(_logFile, "Error outputing " + inPolys + " to " + outPolys);
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
                        FileFunctions.WriteLine(_logFile, "Error outputing " + inFlatTable + " to " + outFlat);
                        return false;
                    }

                    // If metadata .xml file exists delete it.
                    string xmlOutFile = outFlat + ".xml";
                    if (FileFunctions.FileExists(xmlOutFile))
                        FileFunctions.DeleteFile(xmlOutFile); // Not checking for success at the moment.
                }

                FileFunctions.WriteLine(_logFile, "Extract complete");
            }

            // Set the export output format.
            string exportOutputFormat = exportFormat.ToUpper();
            if ((outputType == "CSV") || (outputType == "TXT"))
            {
                FileFunctions.WriteLine(_logFile, "Overriding the export type with '" + outputType + "' ...");
                exportOutputFormat = outputType;
            }

            // Reset the output path.
            outPath = outFolder;

            string expFile;

            // Set the export output name depending on the output format.
            switch (exportOutputFormat)
            {
                case "CSV":
                    if (!string.IsNullOrEmpty(csvFolder))
                        outPath = outFolder + @"\" + csvFolder;
                    expFile = outPath + @"\" + outputTable + ".csv";
                    break;

                case "TXT":
                    if (!string.IsNullOrEmpty(txtFolder))
                        outPath = outFolder + @"\" + txtFolder;
                    expFile = outPath + @"\" + outputTable + ".txt";
                    break;

                default:
                    expFile = outPath + @"\" + outputTable;
                    exportOutputFormat = "";
                    break;
            }

            // Create the output path.
            if (!FileFunctions.DirExists(outPath))
            {
                //FileFunctions.WriteLine(_logFile, "Creating output path '" + outPath + "'.");
                try
                {
                    Directory.CreateDirectory(outPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Cannot create directory " + outPath + ". System error: " + ex.Message);
                    FileFunctions.WriteLine(_logFile, "Cannot create directory " + outPath + ". System error: " + ex.Message);

                    return false;
                }
            }

            // Output the text results if required.
            if (!string.IsNullOrEmpty(exportOutputFormat))
            {
                //TODO - needed?
                //_dockPane.ProgressUpdate(string.Format("Writing output for " + nodeName + " to {0} format ...", exportFormat), -1, 0);

                FileFunctions.WriteLine(_logFile, "Exporting '" + outputTable + "' ...");

                // Check the output geodatabase exists.
                if ((mapOutputFormat == "Geodatabase Table") && (!FileFunctions.DirExists(outPath)))
                {
                    FileFunctions.WriteLine(_logFile, "Creating output geodatabase '" + outPath + "' ...");

                    ArcGISFunctions.CreateFileGeodatabase(outPath);

                    FileFunctions.WriteLine(_logFile, "Output geodatabase created");
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
                            FileFunctions.WriteLine(_logFile, "Error exporting " + inPoints + " to " + outPoints);
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
                            FileFunctions.WriteLine(_logFile, "Error exporting " + inPolys + " to " + outPolys);
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
                        FileFunctions.WriteLine(_logFile, "Error exporting " + inPolys + " to " + outPolys);
                        return false;
                    }

                    // If metadata .xml file exists delete it.
                    string strXmlFile = expFile + ".xml";
                    if (FileFunctions.FileExists(strXmlFile))
                        FileFunctions.DeleteFile(strXmlFile); // Not checking for success at the moment.
                }

                FileFunctions.WriteLine(_logFile, "Export complete");
            }

            // Trigger the macro if one exists
            if (!string.IsNullOrEmpty(macroName))
            {
                FileFunctions.WriteLine(_logFile, "Processing the export file ...");

                if (!StartProcess(macroName, expFile, exportOutputFormat))
                {
                    //MessageBox.Show("Error executing vbscript macro " + mapMacroName + ".");
                    FileFunctions.WriteLine(_logFile, "Error executing vbscript macro " + macroName);
                    _extractErrors = true;

                    return false;
                }
            }

            //TODO - needed?
            //_dockPane.ProgressUpdate("Deleting temporary subset tables", -1, 0);

            // Clean up after processing the layer.
            await ClearSubsetTablesAsync(_defaultSchema, sqlTable, _userID);

            return true;
        }

        private async Task<bool> ProcessMapLayerAsync(Partner partner, MapLayer mapLayer, string outFolder, string partnerFolder,
            string arcGISFolder, string csvFolder, string txtFolder)
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
                FileFunctions.WriteLine(_logFile, "Skipping output = '" + nodeName + "' - not required");
                return true;
            }

            // Does the partner want this layer?
            if (!mapFiles.Contains(nodeName, StringComparison.CurrentCultureIgnoreCase) && !mapFiles.Equals("all", StringComparison.CurrentCultureIgnoreCase))
            {
                FileFunctions.WriteLine(_logFile, "Skipping output = '" + nodeName + "' - not required");
                return true;
            }

            //TODO - needed?
            //_dockPane.ProgressUpdate("Processing Map table " + nodeName + " ...", -1, 0);

            // Replace any date variables in the output name.
            outputTable = outputTable.Replace("%dd%", _dateDD).Replace("%mm%", _dateMM).Replace("%mmm%", _dateMMM).Replace("%mmmm%", _dateMMMM);
            outputTable = outputTable.Replace("%yy%", _dateYY).Replace("%qq%", _dateQQ).Replace("%yyyy%", _dateYYYY).Replace("%ffff%", _dateFFFF);

            // Replace the partner shortname in the output name.
            outputTable = Regex.Replace(outputTable, "%partner%", partnerAbbr, RegexOptions.IgnoreCase);

            //// Set the output file names.
            //string pointFeatureClass = _defaultSchema + "." + mapFiles + "_" + _userID + "_point";
            //string polyFeatureClass = _defaultSchema + "." + mapFiles + "_" + _userID + "_poly";
            //string flatTable = _defaultSchema + "." + mapFiles + "_" + _userID + "_flat";

            //string inPoints = _sdeFileName + @"\" + pointFeatureClass;
            //string inPolys = _sdeFileName + @"\" + polyFeatureClass;
            //string inFlatTable = _sdeFileName + @"\" + flatTable;

            // Build a list of all of the columns required.
            List<string> mapFields = [];
            List<string> rawFields = [.. outputColumns.Split(',')];
            foreach (string mapField in rawFields)
            {
                mapFields.Add(mapField.Trim());
            }

            FileFunctions.WriteLine(_logFile, "");
            FileFunctions.WriteLine(_logFile, "Processing output '" + nodeName + "' ...");

            FileFunctions.WriteLine(_logFile, "Columns names are ... " + outputColumns);
            if (whereClause.Length > 0)
                FileFunctions.WriteLine(_logFile, "Where clause is ... " + whereClause.Replace("\r\n", " "));
            else
                FileFunctions.WriteLine(_logFile, "No where clause is specified.");
            if (orderClause.Length > 0)
                FileFunctions.WriteLine(_logFile, "Order by clause is ... " + orderClause.Replace("\r\n", " "));
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
                FileFunctions.WriteLine(_logFile, "Error creating selection using spatial query");

                _extractErrors = true;
                return false;
            }

            // Find the map layer by name.
            FeatureLayer inputLayer = _mapFunctions.FindLayer(layerName);

            if (inputLayer == null)
                return false;

            // Refine the selection by attributes (if required).
            if (inputLayer.SelectionCount > 0 && !string.IsNullOrEmpty(whereClause))
            {
                FileFunctions.WriteLine(_logFile, "Refining selection with criteria " + whereClause + " ...");

                if (!await _mapFunctions.SelectLayerByAttributesAsync(layerName, whereClause, SelectionCombinationMethod.And))
                {
                    FileFunctions.WriteLine(_logFile, "Error creating subset selection using attribute query");
                    _extractErrors = true;

                    return false;
                }
            }

            // Count the selected features.
            int featureCount = inputLayer.SelectionCount;

            // If there is no selection then exit.
            if (featureCount == 0)
            {
                FileFunctions.WriteLine(_logFile, "There are no features selected in " + layerName + ".");
                return true;
            }

            FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", featureCount) + " records selected.");

            //TODO - needed?
            //_dockPane.ProgressUpdate(string.Format("Exporting selection in layer " + nodeName), -1, 0);

            // Override the output format if an export is required but no output.
            if ((string.IsNullOrEmpty(gisFormat) && !string.IsNullOrEmpty(exportFormat)))
                mapOutputFormat = "SHP";

            // Set the output path.
            string outPath = outFolder;
            if (!string.IsNullOrEmpty(arcGISFolder))
                outPath = outFolder + @"\" + arcGISFolder;

            // Create the output path.
            if (!FileFunctions.DirExists(outPath))
            {
                //FileFunctions.WriteLine(_logFile, "Creating output path '" + outPath + "'.");
                try
                {
                    Directory.CreateDirectory(outPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Cannot create directory " + outPath + ". System error: " + ex.Message);
                    FileFunctions.WriteLine(_logFile, "Cannot create directory " + outPath + ". System error: " + ex.Message);

                    return false;
                }
            }

            // Set the default output file name.
            string outFile = outPath + @"\" + outputTable;

            // Set the map output names depending on the output format.
            switch (mapOutputFormat)
            {
                case "GDB":
                    mapOutputFormat = "Geodatabase FC";
                    outPath = outPath + "\\" + partnerFolder + ".gdb";
                    outFile = outPath + @"\" + outputTable;
                    break;

                case "SHP":
                    mapOutputFormat = "Shapefile";
                    outFile = outPath + @"\" + outputTable + ".shp";

                    break;

                case "DBF":
                    outFile = outPath + @"\" + outputTable + ".dbf";
                    break;

                default:
                    mapOutputFormat = "";
                    break;
            }

            // Output the map results if required.
            if (!string.IsNullOrEmpty(mapOutputFormat))
            {
                //TODO - needed?
                //_dockPane.ProgressUpdate("Writing output for " + nodeName + " to GIS format ...", -1, 0);

                FileFunctions.WriteLine(_logFile, "Extracting '" + outputTable + "' ...");

                // Check the output geodatabase exists.
                if ((mapOutputFormat == "Geodatabase FC") && (!FileFunctions.DirExists(outPath)))
                {
                    FileFunctions.WriteLine(_logFile, "Creating output geodatabase ...");

                    ArcGISFunctions.CreateFileGeodatabase(outPath);

                    FileFunctions.WriteLine(_logFile, "Output geodatabase created");
                }

                // Output the features.
                if (!await ArcGISFunctions.CopyFeaturesAsync(layerName, outFile, true))
                {
                    FileFunctions.WriteLine(_logFile, "Error outputing " + layerName + " to " + outFile);
                    return false;
                }

                // If metadata .xml file exists delete it.
                string xmlOutFile = outFile + ".xml";
                if (FileFunctions.FileExists(xmlOutFile))
                    FileFunctions.DeleteFile(xmlOutFile); // Not checking for success at the moment.

                FileFunctions.WriteLine(_logFile, "Extract complete");
            }

            // Drop non-required fields.
            //FileFunctions.WriteLine(_logFile, "Removing non-required fields from the output file.");
            if (!await _mapFunctions.KeepSelectedFieldsAsync(outputTable, mapFields))
            {
                FileFunctions.WriteLine(_logFile, "Error removing unwanted fields");
                return false;
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

            // Set the default export file name.
            string expFile = outPath + @"\" + outputTable;

            // Set the export output name depending on the output format.
            switch (exportOutputFormat)
            {
                case "CSV":
                    if (!string.IsNullOrEmpty(csvFolder))
                        outPath = outFolder + @"\" + csvFolder;
                    expFile = outPath + @"\" + outputTable + ".csv";
                    break;

                case "TXT":
                    if (!string.IsNullOrEmpty(txtFolder))
                        outPath = outFolder + @"\" + txtFolder;
                    expFile = outPath + @"\" + outputTable + ".txt";
                    break;

                default:
                    exportOutputFormat = "";
                    break;
            }

            // Output the text results if required.
            if (!string.IsNullOrEmpty(exportOutputFormat))
            {
                // Create the output path.
                if (!FileFunctions.DirExists(outPath))
                {
                    //FileFunctions.WriteLine(_logFile, "Creating output path '" + outPath + "'.");
                    try
                    {
                        Directory.CreateDirectory(outPath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Cannot create directory " + outPath + ". System error: " + ex.Message);
                        FileFunctions.WriteLine(_logFile, "Cannot create directory " + outPath + ". System error: " + ex.Message);

                        return false;
                    }
                }

                //TODO - needed?
                //_dockPane.ProgressUpdate(string.Format("Writing output for " + nodeName + " to {0} format ...", exportFormat), -1, 0);

                FileFunctions.WriteLine(_logFile, "Exporting '" + outputTable + "' ...");

                // If schema.ini file exists delete it.
                string strIniFile = FileFunctions.GetDirectoryName(expFile) + "\\schema.ini";
                if (FileFunctions.FileExists(strIniFile))
                    FileFunctions.DeleteFile(strIniFile); // Not checking for success at the moment.

                // Export the features.
                int result;
                if (exportOutputFormat == "TXT")
                    result = await _mapFunctions.CopyFCToTextFileAsync(outputTable, expFile, outputColumns, null, "\t", false, true);
                else
                    result = await _mapFunctions.CopyFCToTextFileAsync(outputTable, expFile, outputColumns, null, ",", false, true);

                if (result <= 0)
                {
                    FileFunctions.WriteLine(_logFile, "Error exporting " + outputTable + " to " + expFile);
                    return false;
                }

                // If metadata .xml file exists delete it.
                string strXmlFile = expFile + ".xml";
                if (FileFunctions.FileExists(strXmlFile))
                    FileFunctions.DeleteFile(strXmlFile); // Not checking for success at the moment.

                FileFunctions.WriteLine(_logFile, "Export complete");
            }

            // Trigger the macro if one exists
            if (!string.IsNullOrEmpty(macroName))
            {
                FileFunctions.WriteLine(_logFile, "Processing the export file ...");

                if (!StartProcess(macroName, expFile, exportOutputFormat))
                {
                    //MessageBox.Show("Error executing vbscript macro " + mapMacroName + ".");
                    FileFunctions.WriteLine(_logFile, "Error executing vbscript macro " + macroName);
                    _extractErrors = true;

                    return false;
                }
            }

            //TODO
            // Remove output layer if not required (i.e. only exported as
            // text file).
            if ((string.IsNullOrEmpty(gisFormat) && !string.IsNullOrEmpty(exportFormat)))
                await _mapFunctions.RemoveLayerAsync(outputTable);

                return true;
        }

        /// <summary>
        /// Perform the spatial selection via a stored procedure.
        /// </summary>
        /// <param name="partner"></param>
        /// <param name="schema"></param>
        /// <param name="selectionTypeNum"></param>
        /// <param name="outputTable"></param>
        /// <param name="useCentroids"></param>
        /// <returns></returns>
        internal async Task<bool> PerformSpatialSelectionAsync(Partner partner, string schema, int selectionTypeNum, string outputTable, bool useCentroids)
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

            // Reset the output counter.
            long tableCount = 0;

            try
            {
                FileFunctions.WriteLine(_logFile, "Executing spatial selection from '" + sqlTable + "' ...");

                // Execute the stored procedure.
                await _sqlFunctions.ExecuteSQLOnGeodatabase(sqlCmd.ToString());

                FileFunctions.WriteLine(_logFile, "Spatial selection complete");

                // Check if the output feature class exists.
                if (!await _sqlFunctions.FeatureClassExistsAsync(outputTable))
                {
                    //TODO - log error message?
                    return false;
                }

                // Count the number of rows in the output feature count.
                tableCount = await _sqlFunctions.FeatureClassCountRowsAsync(outputTable);

                if (tableCount > 0)
                    FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", tableCount) + " records selected");
                else
                    FileFunctions.WriteLine(_logFile, "Procedure returned no records");
            }
            catch (Exception ex)
            {
                FileFunctions.WriteLine(_logFile, "Error executing the stored procedure: " + ex.Message);
                //MessageBox.Show("Error executing the stored procedure: " +
                //    ex.Message, _displayName, MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            return true;
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
                                  string mapOutputFormat)
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

            long maxCount = 0;
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

                        // Count the number of rows in the poly feature count.
                        _polyCount = await _sqlFunctions.FeatureClassCountRowsAsync(polyFeatureClass);

                        // Save the maximum row count.
                        maxCount = _pointCount;
                        if (_pointCount > maxCount)
                            maxCount = _pointCount;

                        if (maxCount == 0)
                        {
                            FileFunctions.WriteLine(_logFile, "No records returned");
                            return false;
                        }

                        FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", _pointCount) + " points to extract");
                        FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", _polyCount) + " polygons to extract");

                        //TODO
                        //// Calculate the total row length for the table.
                        //if (polyCount > 0)
                        //    rowLength = myArcMapFuncs.GetRowLength(strSDEName, strDatabaseSchema + "." + strPolyFC);
                        //else if (pointCount > 0)
                        //    rowLength = intRowLength + myArcMapFuncs.GetRowLength(strSDEName, strDatabaseSchema + "." + strPointFC);
                    }
                    else
                    {
                        // Count the number of rows in the table.
                        _tableCount = await _sqlFunctions.TableCountRowsAsync(flatTable);

                        // Save the maximum row count.
                        maxCount = _tableCount;
                        if (maxCount == 0)
                        {
                            FileFunctions.WriteLine(_logFile, "No records returned");
                            return false;
                        }

                        FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", _tableCount) + " records to extract");
                        FileFunctions.WriteLine(_logFile, "Data is not spatial");

                        //TODO
                        //// Calculate the total row length for the table.
                        //if (tableCount > 0)
                        //    rowLength = myArcMapFuncs.GetRowLength(strSDEName, strDatabaseSchema + "." + strTempTable);
                    }
                }
                else
                {
                    FileFunctions.WriteLine(_logFile, "No records to extract");
                    return false;
                }
            }
            catch (Exception ex)
            {
                FileFunctions.WriteLine(_logFile, "Error executing the stored procedure: " + ex.Message);
                //MessageBox.Show("Error executing the stored procedure: " +
                //    ex.Message, "Data Selector", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            // Check if the maximum record length will be exceeded
            if (rowLength > 4000)
            {
                FileFunctions.WriteLine(_logFile, "Record length exceeds maximum of 4,000 bytes - selection cancelled");
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
                FileFunctions.WriteLine(_logFile, string.Format("Maximum data size = {0} Kb ({1} Gb)", strDataSizeKb, strDataSizeGb));
            }
            else
            {
                if (maxDataSize > 1024)
                    FileFunctions.WriteLine(_logFile, string.Format("Maximum data size = {0} Kb ({1} Mb)", strDataSizeKb, strDataSizeMb));
                else
                    FileFunctions.WriteLine(_logFile, string.Format("Maximum data size = {0} Kb", strDataSizeKb));
            }

            // Check if the maximum data size will be exceeded.
            if ((maxDataSize > (2 * 1024 * 1024)) && (mapOutputFormat == "SHP") || (mapOutputFormat == "DBF"))
            {
                FileFunctions.WriteLine(_logFile, "Maximum data size exceeds 2 Gb - selection cancelled");
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
                FileFunctions.WriteLine(_logFile, "Deleting spatial temporary table");

                // Execute the stored procedure.
                await _sqlFunctions.ExecuteSQLOnGeodatabase(sqlCmd.ToString());
            }
            catch (Exception ex)
            {
                FileFunctions.WriteLine(_logFile, "Error deleting the spatial temporary table: " + ex.Message);
                //MessageBox.Show("Error deleting the temporary tables: " +
                //    ex.Message, "Data Selector", MessageBoxButton.OK, MessageBoxImage.Error);
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
                FileFunctions.WriteLine(_logFile, "Deleting subset temporary tables");

                // Execute the stored procedure.
                await _sqlFunctions.ExecuteSQLOnGeodatabase(sqlCmd.ToString());
            }
            catch (Exception ex)
            {
                FileFunctions.WriteLine(_logFile, "Error deleting the subset temporary tables: " + ex.Message);
                //MessageBox.Show("Error deleting the temporary tables: " +
                //    ex.Message, "Data Selector", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
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
            bool keepSelectedLayers)
        {
            //// Get the settings relevant for this layer.
            //string mapNodeGroup = selectedLayer.NodeGroup;
            ////string mapNodeLayer = selectedLayer.NodeLayer;
            //string mapLayerName = selectedLayer.LayerName;
            //string mapOutputName = selectedLayer.GISOutputName;
            //string mapTableOutputName = selectedLayer.TableOutputName;
            //string mapColumns = selectedLayer.Columns;
            //string mapGroupColumns = selectedLayer.GroupColumns;
            //string mapStatsColumns = selectedLayer.StatisticsColumns;
            //string mapOrderColumns = selectedLayer.OrderColumns;
            //string mapCriteria = selectedLayer.Criteria;

            //bool mapIncludeArea = selectedLayer.IncludeArea;
            //string mapIncludeNearFields = selectedLayer.IncludeNearFields;
            //bool mapIncludeRadius = selectedLayer.IncludeRadius;

            //string mapKeyColumn = selectedLayer.KeyColumn;
            //string mapFormat = selectedLayer.Format;
            //bool mapKeepLayer = selectedLayer.KeepLayer;
            //string mapOutputType = selectedLayer.OutputType;

            //bool mapDisplayLabels = selectedLayer.DisplayLabels;
            //string mapLayerFileName = selectedLayer.LayerFileName;
            //bool mapOverwriteLabels = selectedLayer.OverwriteLabels;
            //string mapLabelColumn = selectedLayer.LabelColumn;
            //string mapLabelClause = selectedLayer.LabelClause;
            //string mapMacroName = selectedLayer.MacroName;

            //string mapCombinedSitesColumns = selectedLayer.CombinedSitesColumns;
            //string mapCombinedSitesGroupColumns = selectedLayer.CombinedSitesGroupColumns;
            //string mapCombinedSitesStatsColumns = selectedLayer.CombinedSitesStatisticsColumns;
            //string mapCombinedSitesOrderColumns = selectedLayer.CombinedSitesOrderByColumns;

            //// Deal with wildcards in the output names.
            //mapOutputName = StringFunctions.ReplaceSearchStrings(mapOutputName, reference, siteName, shortRef, subref, radius);
            //mapTableOutputName = StringFunctions.ReplaceSearchStrings(mapTableOutputName, reference, siteName, shortRef, subref, radius);

            //// Remove any illegal characters from the names.
            //mapOutputName = StringFunctions.StripIllegals(mapOutputName, _repChar);
            //mapTableOutputName = StringFunctions.StripIllegals(mapTableOutputName, _repChar);

            //// Set the statistics columns if they haven't been supplied.
            //if (String.IsNullOrEmpty(mapStatsColumns) && !String.IsNullOrEmpty(mapGroupColumns))
            //    mapStatsColumns = StringFunctions.AlignStatsColumns(mapColumns, mapStatsColumns, mapGroupColumns);
            //if (String.IsNullOrEmpty(mapCombinedSitesStatsColumns) && !String.IsNullOrEmpty(mapCombinedSitesGroupColumns))
            //    mapCombinedSitesStatsColumns = StringFunctions.AlignStatsColumns(mapCombinedSitesColumns, mapCombinedSitesStatsColumns, mapCombinedSitesGroupColumns);

            //// Get the full layer path (in case it's nested in one or more groups).
            //string mapLayerPath = _mapFunctions.GetLayerPath(mapLayerName);

            //// Select by location.
            //FileFunctions.WriteLine(_logFile, "Selecting features using selected feature(s) from layer " + _bufferLayerName + " ...");
            //if (!await ArcGISFunctions.SelectLayerByLocationAsync(mapLayerPath, _bufferLayerPath, "INTERSECT", "", "NEW_SELECTION"))
            //{
            //    //MessageBox.Show("Error selecting layer " + mapLayerName + " by location.");
            //    FileFunctions.WriteLine(_logFile, "Error selecting layer " + mapLayerName + " by location");
            //    _extractErrors = true;

            //    return false;
            //}

            //// Find the map layer by name.
            //FeatureLayer mapLayer = _mapFunctions.FindLayer(mapLayerName);

            //if (mapLayer == null)
            //    return false;

            //// Refine the selection by attributes (if required).
            //if (mapLayer.SelectionCount > 0 && !string.IsNullOrEmpty(mapCriteria))
            //{
            //    FileFunctions.WriteLine(_logFile, "Refining selection with criteria " + mapCriteria + " ...");

            //    if (!await _mapFunctions.SelectLayerByAttributesAsync(mapLayerName, mapCriteria, SelectionCombinationMethod.And))
            //    {
            //        //MessageBox.Show("Error selecting layer " + mapLayerName + " with criteria " + mapCriteria + ". Please check syntax and column names (case sensitive).");
            //        FileFunctions.WriteLine(_logFile, "Error refining selection on layer " + mapLayerName + " with criteria " + mapCriteria + ". Please check syntax and column names (case sensitive)");
            //        _extractErrors = true;

            //        return false;
            //    }
            //}

            //// Count the selected features.
            //int featureCount = mapLayer.SelectionCount;

            //// Write out the results - to a feature class initially. Include distance if required.
            //if (featureCount <= 0)
            //{
            //    FileFunctions.WriteLine(_logFile, "No features found");
            //    return true;
            //}

            //FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", featureCount) + " feature(s) found");

            //// Create the map output depending on the output type required.
            //if (!await CreateMapOutputAsync(mapLayerName, mapLayerPath, _bufferLayerPath, mapOutputType))
            //{
            //    MessageBox.Show("Cannot output selection from " + mapLayerName + " to " + _tempMasterOutputFile + ".");
            //    FileFunctions.WriteLine(_logFile, "Cannot output selection from " + mapLayerName + " to " + _tempMasterOutputFile);

            //    return false;
            //}

            //// Add map labels to the output if required.
            //if (addSelectedLayersOption == AddSelectedLayersOptions.WithLabels && !string.IsNullOrEmpty(mapLabelColumn))
            //{
            //    if (!await AddMapLabelsAsync(overwriteLabelOption, mapOverwriteLabels, mapLabelColumn, mapKeyColumn, mapNodeGroup))
            //    {
            //        //MessageBox.Show("Error adding map labels to " + mapLabelColumn + " in " + _tempMasterOutputFile + ".");
            //        FileFunctions.WriteLine(_logFile, "Error adding map labels to " + mapLabelColumn + " in " + _tempMasterOutputFile);
            //        _extractErrors = true;

            //        return false;
            //    }
            //}

            //// Create relevant output names.
            //string mapOutputFile = _outputFolder + @"\" + mapOutputName; // Output shapefile / feature class name. Note no extension to allow write to GDB.
            //string mapTableOutputFile = _outputFolder + @"\" + mapTableOutputName + "." + mapFormat.ToLower(); // Output table name.

            //// Include headers for CSV files.
            //bool includeHeaders = false;
            //if (mapFormat.Equals("csv", StringComparison.OrdinalIgnoreCase))
            //    includeHeaders = true;

            //// Only include radius if requested.
            //string radiusText = "none";
            //if (mapIncludeRadius)
            //    radiusText = radius;

            //string areaUnit = "";
            //if (mapIncludeArea)
            //    areaUnit = areaMeasureUnit;

            //// Export results to table if required.
            //if (!string.IsNullOrEmpty(mapFormat) && !string.IsNullOrEmpty(mapColumns))
            //{
            //    FileFunctions.WriteLine(_logFile, "Extracting summary information ...");

            //    int intLineCount = await ExportSelectionAsync(mapTableOutputFile, mapFormat.ToLower(), mapColumns, mapGroupColumns, mapStatsColumns, mapOrderColumns,
            //        includeHeaders, false, areaUnit, mapIncludeNearFields, radiusText);
            //    if (intLineCount <= 0)
            //    {
            //        //MessageBox.Show("Error extracting summary from " + _tempMasterOutputFile + ".");
            //        FileFunctions.WriteLine(_logFile, "Error extracting summary from " + _tempMasterOutputFile);
            //        _extractErrors = true;

            //        return false;
            //    }

            //    FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", intLineCount) + " record(s) exported");
            //}

            //// If selected layers are to be kept, and this layer is to be kept,
            //// copy to a permanent layer.
            //if ((keepSelectedLayers) && (mapKeepLayer))
            //{
            //    if (!await KeepLayerAsync(mapOutputName, mapOutputFile, addSelectedLayersOption, mapLayerFileName, mapDisplayLabels, mapLabelClause, mapLabelColumn))
            //    {
            //        _extractErrors = true;

            //        return false;
            //    }
            //}

            //// Add to combined sites table if required.
            //if (!string.IsNullOrEmpty(mapCombinedSitesColumns) && combinedSitesTableOption != CombinedSitesTableOptions.None)
            //{
            //    FileFunctions.WriteLine(_logFile, "Extracting summary output for combined sites table ...");

            //    int intLineCount = await ExportSelectionAsync(_combinedSitesOutputFile, _combinedSitesTableFormat, mapCombinedSitesColumns, mapCombinedSitesGroupColumns,
            //        mapCombinedSitesStatsColumns, mapCombinedSitesOrderColumns,
            //        false, true, areaUnit, mapIncludeNearFields, radiusText);

            //    if (intLineCount < 0)
            //    {
            //        //MessageBox.Show("Error extracting summary for combined sites table from " + _tempMasterOutputFile + ".");
            //        FileFunctions.WriteLine(_logFile, "Error extracting summary for combined sites table from " + _tempMasterOutputFile);
            //        _extractErrors = true;

            //        return false;
            //    }

            //    FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", intLineCount) + " row(s) added to combined sites table");
            //}

            //// Cleanup the temporary master layer.
            ////await _mapFunctions.RemoveLayerAsync(_tempMasterLayerName);
            ////await ArcGISFunctions.DeleteFeatureClassAsync(_tempMasterOutputFile);

            //// Clear the selection in the input layer.
            //await _mapFunctions.ClearLayerSelectionAsync(mapLayerName);

            //FileFunctions.WriteLine(_logFile, "Analysis complete");

            //// Trigger the macro if one exists
            //if (!string.IsNullOrEmpty(mapMacroName))
            //{
            //    FileFunctions.WriteLine(_logFile, "Executing vbscript macro ...");

            //    if (!StartProcess(mapMacroName, mapTableOutputName, mapFormat))
            //    {
            //        //MessageBox.Show("Error executing vbscript macro " + mapMacroName + ".");
            //        FileFunctions.WriteLine(_logFile, "Error executing vbscript macro " + mapMacroName);
            //        _extractErrors = true;

            //        return false;
            //    }
            //}

            return true;
        }

        /// <summary>
        /// Create the required output type from the current layer.
        /// </summary>
        /// <param name="mapLayerName"></param>
        /// <param name="mapLayerPath"></param>
        /// <param name="PartnerLayerPath"></param>
        /// <param name="mapOutputType"></param>
        /// <returns></returns>
        private async Task<bool> CreateMapOutputAsync(string mapLayerName, string mapLayerPath, string PartnerLayerPath, string mapOutputType)
        {
            //TODO
            return false;

            //// Get the input feature class type.
            //string mapLayerFCType = _mapFunctions.GetFeatureClassType(mapLayerName);
            //if (mapLayerFCType == null)
            //    return false;

            //// Get the buffer feature class type.
            //string bufferFCType = _mapFunctions.GetFeatureClassType(_bufferLayerName);
            //if (bufferFCType == null)
            //    return false;

            //// If the input layer should be clipped to the buffer layer, do so now.
            //if (mapOutputType.Equals("CLIP", StringComparison.OrdinalIgnoreCase))
            //{
            //    if (mapLayerFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase) && bufferFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase) ||
            //        mapLayerFCType.Equals("line", StringComparison.OrdinalIgnoreCase) &&
            //        (bufferFCType.Equals("line", StringComparison.OrdinalIgnoreCase) || bufferFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase)))
            //    {
            //        // Clip
            //        FileFunctions.WriteLine(_logFile, "Clipping selected features ...");
            //        return await ArcGISFunctions.ClipFeaturesAsync(mapLayerPath, PartnerLayerPath, _tempMasterOutputFile, true);
            //    }
            //    else
            //    {
            //        // Copy
            //        FileFunctions.WriteLine(_logFile, "Copying selected features ...");
            //        return await ArcGISFunctions.CopyFeaturesAsync(mapLayerPath, _tempMasterOutputFile, true);
            //    }
            //}
            //// If the buffer layer should be clipped to the input layer, do so now.
            //else if (mapOutputType.Equals("OVERLAY", StringComparison.OrdinalIgnoreCase))
            //{
            //    if (bufferFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase) && mapLayerFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase) ||
            //        bufferFCType.Equals("line", StringComparison.OrdinalIgnoreCase) &&
            //        (mapLayerFCType.Equals("line", StringComparison.OrdinalIgnoreCase) || mapLayerFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase)))
            //    {
            //        // Clip
            //        FileFunctions.WriteLine(_logFile, "Overlaying selected features ...");
            //        return await ArcGISFunctions.ClipFeaturesAsync(PartnerLayerPath, mapLayerPath, _tempMasterOutputFile, true);
            //    }
            //    else
            //    {
            //        // Select from the buffer layer.
            //        FileFunctions.WriteLine(_logFile, "Selecting features  ...");
            //        await ArcGISFunctions.SelectLayerByLocationAsync(PartnerLayerPath, mapLayerPath);

            //        // Find the buffer layer by name.
            //        FeatureLayer bufferLayer = _mapFunctions.FindLayer(_bufferLayerName);

            //        if (bufferLayer == null)
            //            return false;

            //        // Count the selected features.
            //        int featureCount = bufferLayer.SelectionCount;
            //        if (featureCount > 0)
            //        {
            //            FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", featureCount) + " feature(s) found");

            //            // Copy the selection from the buffer layer.
            //            FileFunctions.WriteLine(_logFile, "Copying selected features ... ");
            //            if (!await ArcGISFunctions.CopyFeaturesAsync(PartnerLayerPath, _tempMasterOutputFile, true))
            //                return false;
            //        }
            //        else
            //        {
            //            FileFunctions.WriteLine(_logFile, "No features selected");

            //            return true;
            //        }

            //        // Clear the buffer layer selection.
            //        await _mapFunctions.ClearLayerSelectionAsync(_bufferLayerName);

            //        return true;
            //    }
            //}
            //// If the input layer should be intersected with the buffer layer, do so now.
            //else if (mapOutputType.Equals("INTERSECT", StringComparison.OrdinalIgnoreCase))
            //{
            //    if (mapLayerFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase) && bufferFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase) ||
            //        mapLayerFCType.Equals("line", StringComparison.OrdinalIgnoreCase) &&
            //        (bufferFCType.Equals("line", StringComparison.OrdinalIgnoreCase) || bufferFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase)))
            //    {
            //        string[] features = ["'" + mapLayerPath + "' #", "'" + PartnerLayerPath + "' #"];
            //        string inFeatures = string.Join(";", features);

            //        // Intersect
            //        FileFunctions.WriteLine(_logFile, "Intersecting selected features ...");
            //        return await ArcGISFunctions.IntersectFeaturesAsync(inFeatures, _tempMasterOutputFile, addToMap: true); // Selected features in input, buffer FC, output.
            //    }
            //    else
            //    {
            //        // Copy
            //        FileFunctions.WriteLine(_logFile, "Copying selected features ...");
            //        return await ArcGISFunctions.CopyFeaturesAsync(mapLayerPath, _tempMasterOutputFile, true);
            //    }
            //}
            //// Otherwise do a straight copy of the input layer.
            //else
            //{
            //    // Copy
            //    FileFunctions.WriteLine(_logFile, "Copying selected features ...");
            //    return await ArcGISFunctions.CopyFeaturesAsync(mapLayerPath, _tempMasterOutputFile, true);
            //}
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
            int intLineCount = 0;

            //// Only export if the user has specified columns.
            //if (string.IsNullOrEmpty(mapColumns))
            //    return -1;

            //// Check the input feature layer exists.
            //FeatureLayer inputFeaturelayer = _mapFunctions.FindLayer(_tempMasterLayerName);
            //if (inputFeaturelayer == null)
            //    return -1;

            //// Get the input feature class type.
            //string inputFeatureType = _mapFunctions.GetFeatureClassType(inputFeaturelayer);
            //if (inputFeatureType == null)
            //    return -1;

            //// Calculate the area field if required.
            //string areaColumnName = "";
            //if (!string.IsNullOrEmpty(areaUnit) && inputFeatureType.Equals("polygon", StringComparison.OrdinalIgnoreCase))
            //{
            //    areaColumnName = "Area" + areaUnit;
            //    // Does the area field already exist? If not, add it.
            //    if (!await _mapFunctions.FieldExistsAsync(_tempMasterLayerName, areaColumnName))
            //    {
            //        if (!await ArcGISFunctions.AddFieldAsync(_tempMasterOutputFile, areaColumnName, "DOUBLE", 20))
            //        {
            //            //MessageBox.Show("Error adding area field to " + _tempMasterOutputFile + ".");
            //            FileFunctions.WriteLine(_logFile, "Error adding area field to " + _tempMasterOutputFile);
            //            _extractErrors = true;

            //            return -1;
            //        }

            //        string geometryProperty = areaColumnName + " AREA";
            //        if (areaUnit.Equals("ha", StringComparison.OrdinalIgnoreCase))
            //        {
            //            areaUnit = "HECTARES";
            //        }
            //        else if (areaUnit.Equals("m2", StringComparison.OrdinalIgnoreCase))
            //        {
            //            areaUnit = "SQUARE_METERS";
            //        }
            //        else if (areaUnit.Equals("km2", StringComparison.OrdinalIgnoreCase))
            //        {
            //            areaUnit = "SQUARE_KILOMETERS";
            //        }

            //        // Calculate the area field.
            //        if (!await ArcGISFunctions.CalculateGeometryAsync(_tempMasterOutputFile, geometryProperty, "", areaUnit))
            //        {
            //            //MessageBox.Show("Error calculating area field in " + _tempMasterOutputFile + ".");
            //            FileFunctions.WriteLine(_logFile, "Error calculating area field in " + _tempMasterOutputFile);
            //            _extractErrors = true;

            //            return -1;
            //        }
            //    }
            //}

            //// Include radius if requested
            //if (radiusText != "none")
            //{
            //    FileFunctions.WriteLine(_logFile, "Including radius column ...");

            //    // Does the radius field already exist? If not, add it.
            //    if (!await _mapFunctions.FieldExistsAsync(_tempMasterLayerName, "Radius"))
            //    {
            //        if (!await ArcGISFunctions.AddFieldAsync(_tempMasterOutputFile, "Radius", "TEXT", fieldLength: 25))
            //        {
            //            //MessageBox.Show("Error adding radius field to " + _tempMasterOutputFile + ".");
            //            FileFunctions.WriteLine(_logFile, "Error adding radius field to " + _tempMasterOutputFile);
            //            _extractErrors = true;

            //            return -1;
            //        }
            //    }

            //    // Calculate the radius field.
            //    if (!await ArcGISFunctions.CalculateFieldAsync(_tempMasterOutputFile, "Radius", '"' + radiusText + '"'))
            //    {
            //        //MessageBox.Show("Error calculating radius field in " + _tempMasterOutputFile + ".");
            //        FileFunctions.WriteLine(_logFile, "Error calculating radius field in " + _tempMasterOutputFile);
            //        _extractErrors = true;

            //        return -1;
            //    }
            //}

            //// Copy the input features.
            //if (!await ArcGISFunctions.CopyFeaturesAsync(_tempMasterOutputFile, _tempFCOutputFile, true))
            //{
            //    //MessageBox.Show("Error copying output file to " + _tempFCOutputFile + ".");
            //    FileFunctions.WriteLine(_logFile, "Error copying output file to " + _tempFCOutputFile);
            //    _extractErrors = true;

            //    return -1;
            //}

            ////-------------------------------------------------------------
            //// After this the input to the remainder of the function
            //// should be reading from _tempFCOutputFile (_tempFCLayerName).
            ////-------------------------------------------------------------

            //// Calculate the boundary distance and bearing if required.
            //if (includeNearFields.Equals("BOUNDARY", StringComparison.OrdinalIgnoreCase))
            //{
            //    // Calculate the distance and additional proximity fields.
            //    if (!await ArcGISFunctions.NearAnalysisAsync(_tempFCOutputFile, _searchLayerName,
            //        radiusText, "LOCATION", "ANGLE", "PLANAR", null, "METERS"))
            //    {
            //        //MessageBox.Show("Error calculating nearest distance from " + _tempFCPointsOutputFile + " to " + _tempSearchPointsOutputFile + ".");
            //        FileFunctions.WriteLine(_logFile, "Error calculating nearest distance from " + _tempFCOutputFile + " to " + _searchLayerName);
            //        _extractErrors = true;

            //        return -1;
            //    }
            //}
            //// Calculate the centroid distance and bearing if required.
            //else if (includeNearFields.Equals("CENTROID", StringComparison.OrdinalIgnoreCase))
            //{
            //    // Convert the output features to points.
            //    if (!await ArcGISFunctions.FeatureToPointAsync(_tempFCOutputFile, _tempFCPointsOutputFile,
            //        "CENTROID", addToMap: false))
            //    {
            //        //MessageBox.Show("Error converting " + _tempFCOutputFile + " features to points into " + _tempFCPointsOutputFile + ".");
            //        FileFunctions.WriteLine(_logFile, "Error converting " + _tempFCOutputFile + " features to points into " + _tempFCPointsOutputFile);
            //        _extractErrors = true;

            //        return -1;
            //    }

            //    // Convert the search features to points.
            //    if (!await ArcGISFunctions.FeatureToPointAsync(_searchLayerName, _tempSearchPointsOutputFile,
            //        "CENTROID", addToMap: false))
            //    {
            //        //MessageBox.Show("Error converting " + _searchLayerName + " features to points into " + _tempSearchPointsOutputFile + ".");
            //        FileFunctions.WriteLine(_logFile, "Error converting " + _searchLayerName + " features to points into " + _tempSearchPointsOutputFile);
            //        _extractErrors = true;

            //        return -1;
            //    }

            //    // Calculate the distance and additional proximity fields.
            //    if (!await ArcGISFunctions.NearAnalysisAsync(_tempFCPointsOutputFile, _tempSearchPointsOutputFile,
            //        radiusText, "LOCATION", "ANGLE", "PLANAR", null, "METERS"))
            //    {
            //        //MessageBox.Show("Error calculating nearest distance from " + _tempFCPointsOutputFile + " to " + _tempSearchPointsOutputFile + ".");
            //        FileFunctions.WriteLine(_logFile, "Error calculating nearest distance from " + _tempFCPointsOutputFile + " to " + _tempSearchPointsOutputFile);
            //        _extractErrors = true;

            //        return -1;
            //    }

            //    string joinFields = "NEAR_DIST;NEAR_ANGLE";

            //    // Join the distance and addition proximity fields to the output feature layer.
            //    if (!await ArcGISFunctions.JoinFieldsAsync(_tempFCLayerName, "OBJECTID", _tempFCPointsOutputFile, "ORIG_FID",
            //        joinFields, addToMap: true))
            //    {
            //        //MessageBox.Show("Error joining fields to " + _tempFCLayerName + " from " + _tempFCPointsOutputFile + ".");
            //        FileFunctions.WriteLine(_logFile, "Error joining fields to " + _tempFCLayerName + " from " + _tempFCPointsOutputFile);
            //        _extractErrors = true;

            //        return -1;
            //    }
            //}

            //// Check the output feature layer exists.
            //FeatureLayer outputFeatureLayer = _mapFunctions.FindLayer(_tempFCLayerName);
            //if (outputFeatureLayer == null)
            //    return -1;

            //// Check all the requested group by fields exist.
            //// Only pass those that do.
            //if (!string.IsNullOrEmpty(mapGroupColumns))
            //{
            //    List<string> groupColumnList = [.. mapGroupColumns.Split(';')];
            //    mapGroupColumns = "";
            //    foreach (string groupColumn in groupColumnList)
            //    {
            //        string columnName = groupColumn.Trim();

            //        if (await _mapFunctions.FieldExistsAsync(_tempFCLayerName, columnName))
            //            mapGroupColumns = mapGroupColumns + columnName + ";";
            //    }
            //    if (!string.IsNullOrEmpty(mapGroupColumns))
            //        mapGroupColumns = mapGroupColumns.Substring(0, mapGroupColumns.Length - 1);
            //}

            //// Check all the requested statistics fields exist.
            //// Only pass those that do.
            //if (!string.IsNullOrEmpty(mapStatsColumns))
            //{
            //    List<string> statsColumnList = [.. mapStatsColumns.Split(';')];
            //    mapStatsColumns = "";
            //    foreach (string statsColumn in statsColumnList)
            //    {
            //        List<string> statsComponents = [.. statsColumn.Split(' ')];
            //        string columnName = statsComponents[0].Trim(); // The field name.

            //        if (await _mapFunctions.FieldExistsAsync(_tempFCLayerName, columnName))
            //            mapStatsColumns = mapStatsColumns + statsColumn + ";";
            //    }
            //    if (!string.IsNullOrEmpty(mapStatsColumns))
            //        mapStatsColumns = mapStatsColumns.Substring(0, mapStatsColumns.Length - 1);
            //}

            //// If we have group columns but no statistics columns, add a dummy column.
            //if (string.IsNullOrEmpty(mapStatsColumns) && !string.IsNullOrEmpty(mapGroupColumns))
            //{
            //    string strDummyField = mapGroupColumns.Split(';').ToList()[0];
            //    mapStatsColumns = strDummyField + " FIRST";
            //}

            //// Now do the summary statistics as required, or export the layer to table if not.
            //if (!string.IsNullOrEmpty(mapStatsColumns))
            //{
            //    FileFunctions.WriteLine(_logFile, "Calculating summary statistics ...");

            //    string statisticsFields = "";
            //    if (!string.IsNullOrEmpty(mapStatsColumns))
            //        statisticsFields = mapStatsColumns;

            //    string caseFields = "";
            //    if (!string.IsNullOrEmpty(mapGroupColumns))
            //        caseFields = mapGroupColumns;

            //    // Add the radius column to the stats columns if it's not already there.
            //    if (radiusText != "none")
            //    {
            //        if (!statisticsFields.Contains("Radius FIRST", StringComparison.OrdinalIgnoreCase))
            //            statisticsFields += ";Radius FIRST";
            //    }

            //    // Add the area column to the stats columns if it's not already there.
            //    if (!string.IsNullOrEmpty(areaColumnName))
            //    {
            //        if (!statisticsFields.Contains(areaColumnName + " FIRST", StringComparison.OrdinalIgnoreCase))
            //            statisticsFields += ";" + areaColumnName + " FIRST";
            //    }

            //    // Calculate the summary statistics.
            //    if (!await ArcGISFunctions.CalculateSummaryStatisticsAsync(_tempFCOutputFile, _tempTableOutputFile, statisticsFields, caseFields, addToMap: true))
            //    {
            //        //MessageBox.Show("Error calculating summary statistics for '" + _tempFCOutputFile + "' into " + _tempTableOutputFile + ".");
            //        FileFunctions.WriteLine(_logFile, "Error calculating summary statistics for '" + _tempFCOutputFile + "' into " + _tempTableOutputFile);
            //        _extractErrors = true;

            //        return -1;
            //    }

            //    // Get the list of fields for the input table.
            //    IReadOnlyList<Field> inputFields;
            //    inputFields = await _mapFunctions.GetTableFieldsAsync(_tempTableLayerName);

            //    // Check a list of fields is returned.
            //    if (inputFields == null || inputFields.Count == 0)
            //        return -1;

            //    // Now rename the radius field.
            //    if (radiusText != "none")
            //    {
            //        string oldFieldName;
            //        // Check the radius field by name.
            //        try
            //        {
            //            oldFieldName = inputFields.Where(f => f.Name.Equals("FIRST_Radius", StringComparison.OrdinalIgnoreCase)).First().Name;
            //        }
            //        catch
            //        {
            //            // If not found then use the last field.
            //            int intNewIndex = inputFields.Count - 1;
            //            oldFieldName = inputFields[intNewIndex].Name;
            //        }

            //        if (!await ArcGISFunctions.RenameFieldAsync(_tempTableOutputFile, oldFieldName, "Radius"))
            //        {
            //            //MessageBox.Show("Error renaming radius field in " + _tempFCOutputFile + ".");
            //            FileFunctions.WriteLine(_logFile, "Error renaming radius field in " + _tempTableLayerName);
            //            _extractErrors = true;

            //            return -1;
            //        }
            //    }

            //    // Now rename the area field.
            //    if (!string.IsNullOrEmpty(areaColumnName))
            //    {
            //        string oldFieldName;
            //        // Check the area field by name.
            //        try
            //        {
            //            oldFieldName = inputFields.Where(f => f.Name.Equals("FIRST_" + areaColumnName, StringComparison.OrdinalIgnoreCase)).First().Name;
            //        }
            //        catch
            //        {
            //            // If not found then use the last field.
            //            int intNewIndex = inputFields.Count - 1;
            //            oldFieldName = inputFields[intNewIndex].Name;
            //        }

            //        if (!await ArcGISFunctions.RenameFieldAsync(_tempTableOutputFile, oldFieldName, areaColumnName))
            //        {
            //            //MessageBox.Show("Error renaming Area field in " + _tempFCOutputFile + ".");
            //            FileFunctions.WriteLine(_logFile, "Error renaming Area field in " + _tempTableLayerName);
            //            _extractErrors = true;

            //            return -1;
            //        }
            //    }

            //    // Now export the output table.
            //    FileFunctions.WriteLine(_logFile, "Exporting to " + outputFormat.ToUpper() + " ...");
            //    intLineCount = await _mapFunctions.CopyTableToTextFileAsync(_tempTableLayerName, outputTableName, mapColumns, mapOrderColumns, append, includeHeaders);
            //}
            //else
            //{
            //    // Do straight copy of the feature class.
            //    FileFunctions.WriteLine(_logFile, "Exporting to " + outputFormat.ToUpper() + " ...");
            //    intLineCount = await _mapFunctions.CopyFCToTextFileAsync(_tempFCLayerName, outputTableName, mapColumns, mapOrderColumns, append, includeHeaders);
            //}

            return intLineCount;
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

            //TODO
            //scriptProc.StartInfo.FileName = @"cscript.exe";
            //scriptProc.StartInfo.WorkingDirectory = FileFunctions.GetDirectoryName(macroName); //<---very important
            //scriptProc.StartInfo.UseShellExecute = true;
            //scriptProc.StartInfo.Arguments = string.Format(@"//B //Nologo {0} {1} {2} {3}", "\"" + macroName + "\"", "\"" + _outputFolder + "\"", "\"" + mapTableOutputName + "." + mapFormat.ToLower() + "\"", "\"" + mapTableOutputName + ".xlsx" + "\"");
            //scriptProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden; //prevent console window from popping up

            try
            {
                scriptProc.Start();
                scriptProc.WaitForExit(); // <-- Optional if you want program running until your script exits.

                int exitcode = scriptProc.ExitCode;
                if (exitcode != 0)
                {
                    FileFunctions.WriteLine(_logFile, "Error executing vbscript macro. Exit code : " + exitcode);
                    _extractErrors = true;

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
                    Action<object> clearAction = new(PartnersListExpandCommandClick);
                    _partnersListExpandCommand = new RelayCommand(clearAction, param => true);
                }
                return _partnersListExpandCommand;
            }
        }

        /// <summary>
        /// Handles event when Cancel button is pressed.
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
        /// Create PartnersList Expand button command.
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
                    Action<object> clearAction = new(SQLLayersListExpandCommandClick);
                    _sqlLayersListExpandCommand = new RelayCommand(clearAction, param => true);
                }
                return _sqlLayersListExpandCommand;
            }
        }

        /// <summary>
        /// Handles event when Cancel button is pressed.
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
        /// Create PartnersList Expand button command.
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
                    Action<object> clearAction = new(MapLayersListExpandCommandClick);
                    _mapLayersListExpandCommand = new RelayCommand(clearAction, param => true);
                }
                return _mapLayersListExpandCommand;
            }
        }

        /// <summary>
        /// Handles event when Cancel button is pressed.
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
        /// <param name="refresh"></param>
        /// <returns></returns>
        public async Task GetSQLTableNamesAsync(bool refresh)
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

        public bool IsActive { get; set; }

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
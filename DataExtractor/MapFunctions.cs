// The DataTools are a suite of ArcGIS Pro addins used to extract, sync
// and manage biodiversity information from ArcGIS Pro and SQL Server
// based on pre-defined or user specified criteria.
//
// Copyright © 2024-25 Andy Foy Consulting.
//
// This file is part of DataTools suite of programs.
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

using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Data.DDL;
using ArcGIS.Core.Data.Exceptions;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Catalog;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Geoprocessing;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Editing.Attributes;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Internal.Layouts.Utilities;
using ArcGIS.Desktop.Layouts;
using ArcGIS.Desktop.Mapping;
using DataExtractor;
using DataExtractor.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Threading.Tasks;
using QueryFilter = ArcGIS.Core.Data.QueryFilter;

namespace DataTools
{
    /// <summary>
    /// This class provides ArcGIS Pro map functions.
    /// </summary>
    internal class MapFunctions
    {
        #region Fields

        private Map _activeMap;

        #endregion Fields

        #region Constructor

        /// <summary>
        /// Set the global variables.
        /// </summary>
        public MapFunctions()
        {
            // Get the active map view (if there is one).
            MapView activeMapView = GetActiveMapView();

            // Set the map currently displayed in the active map view.
            if (activeMapView != null)
                _activeMap = activeMapView.Map;
            else
                _activeMap = null;
        }

        #endregion Constructor

        #region Properties

        /// <summary>
        /// The name of the active map.
        /// </summary>
        public string MapName
        {
            get
            {
                // If there is no active map, return null.
                if (_activeMap == null)
                    return null;
                else
                    return _activeMap.Name;
            }
        }

        #endregion Properties

        #region Debug Logging

        private const string ToolLogPrefix = "DE|";

        /// <summary>
        /// Writes any message to the Trace log with a timestamp.
        /// </summary>
        /// <param name="message">The message to log.</param>
        private static void TraceLog(string message)
        {
            Trace.WriteLine($"{ToolLogPrefix}{DateTime.Now:G} : {message}");
        }

        #endregion Debug Logging

        #region Map

        /// <summary>
        /// Retrieves the currently active map view, if one is available.
        /// </summary>
        /// <returns>
        /// The active <see cref="MapView"/> instance, or <c>null</c> if no map view is active.
        /// </returns>
        internal static MapView GetActiveMapView()
        {
            // Get the active map view from the ArcGIS Pro application.
            MapView mapView = MapView.Active;

            // Return the map view if available; otherwise, return null.
            return mapView;
        }

        /// <summary>
        /// Retrieves a map from the current project by its name.
        /// </summary>
        /// <param name="mapName">The name of the map to retrieve.</param>
        /// <returns>
        /// A <see cref="Map"/> instance if found; otherwise, <c>null</c>.
        /// </returns>
        public async Task<Map> GetMapFromNameAsync(string mapName)
        {
            // Return null if the input name is invalid.
            if (mapName == null)
            {
                TraceLog("GetMapFromNameAsync error: No map name provided.");
                return null;
            }

            Map map = null;

            // Run on the CIM thread to access project items safely.
            await QueuedTask.Run(() =>
            {
                // Search for the map project item by name and retrieve the associated Map object.
                map = Project.Current.GetItems<MapProjectItem>()
                    .FirstOrDefault(m => m.Name == mapName)
                    ?.GetMap();
            });

            // Return the found map, or null if not found.
            return map;
        }

        /// <summary>
        /// Resolves the <see cref="Map"/> associated with a map pane using its caption, activating the pane if needed.
        /// </summary>
        /// <param name="mapViewCaption">The caption of the map pane (tab title in ArcGIS Pro).</param>
        /// <returns>
        /// The corresponding <see cref="Map"/> if the pane is open and initialized; otherwise, <c>null</c>.
        /// </returns>
        public async Task<Map> GetMapFromCaptionAsync(string mapViewCaption)
        {
            if (string.IsNullOrWhiteSpace(mapViewCaption))
            {
                TraceLog("GetMapFromCaptionAsync error: No caption provided.");
                return null;
            }

            // Find the map pane by caption (regardless of activation state).
            var pane = FrameworkApplication.Panes
                .OfType<IMapPane>()
                .FirstOrDefault(p => (p as Pane)?.Caption.Equals(mapViewCaption.Trim(), StringComparison.OrdinalIgnoreCase) == true);

            if (pane == null)
            {
                TraceLog($"GetMapFromCaptionAsync error: No map pane found for caption '{mapViewCaption}'.");
                return null;
            }

            // Activate the pane to ensure MapView is fully initialized.
            (pane as Pane)?.Activate();

            // Retry loop: wait for MapView?.Map to be non-null (up to 5 seconds).
            const int maxWaitMs = 5000;
            const int delayIntervalMs = 200;
            int elapsedMs = 0;

            while (elapsedMs < maxWaitMs)
            {
                var map = pane.MapView?.Map;
                if (map != null)
                {
                    return map;
                }

                await Task.Delay(delayIntervalMs);
                elapsedMs += delayIntervalMs;
            }

            TraceLog($"GetMapFromCaptionAsync error: MapView is still null after waiting {maxWaitMs}ms for pane '{mapViewCaption}'.");

            return null;
        }

        /// <summary>
        /// Opens the specified map in a new pane if it's not already open, and activates it.
        /// </summary>
        /// <param name="map">The Map object to activate or open.</param>
        internal static async Task OpenMapAsync(Map map)
        {
            // Check if a pane is already open for this map.
            var pane = ProApp.Panes
                .OfType<Pane>()
                .FirstOrDefault(p =>
                {
                    if (p is IMapPane mp && mp.MapView.Map == map)
                        return true;
                    return false;
                });

            if (pane != null)
            {
                // Already open — activate it.
                pane.Activate();

                // Check it worked.
                var isActive = MapView.Active?.Map == map;
            }
            else
            {
                // Not open — open and activate it.
                var newPane = await ProApp.Panes.CreateMapPaneAsync(map);
            }
        }

        /// <summary>
        /// Activates the pane displaying the specified <see cref="Map"/> and returns its associated <see cref="MapView"/>.
        /// Falls back to the internally stored active map if no map is provided.
        /// </summary>
        /// <param name="targetMap">The map to activate, or <c>null</c> to use the internally stored active map.</param>
        /// <returns>
        /// The <see cref="MapView"/> associated with the activated pane, or <c>null</c> if not found.
        /// </returns>
        public async Task<MapView> ActivateMapAsync(Map targetMap)
        {
            // Use the provided map or fall back to the internally stored active map.
            Map mapToUse = targetMap ?? _activeMap;

            if (mapToUse == null)
            {
                TraceLog("ActivateMapAsync error: No map provided and no fallback map available.");
                return null;
            }

            // Search for an open map pane whose MapView references the target map.
            var pane = FrameworkApplication.Panes
                .OfType<IMapPane>()
                .FirstOrDefault(p => p.MapView?.Map == mapToUse);

            if (pane == null)
            {
                TraceLog($"ActivateMapAsync error: No open pane found for map '{mapToUse.Name}'.");
                return null;
            }

            // Activate the pane.
            (pane as Pane)?.Activate();

            // Retry loop: wait for MapView to be non-null (up to 5 seconds).
            const int maxWaitMs = 5000;
            const int delayIntervalMs = 200;
            int elapsedMs = 0;

            while (elapsedMs < maxWaitMs)
            {
                var mapView = pane.MapView;
                if (mapView != null)
                    return mapView;

                await Task.Delay(delayIntervalMs);
                elapsedMs += delayIntervalMs;
            }

            TraceLog($"ActivateMapAsync error: MapView is still null after waiting {maxWaitMs}ms for map '{mapToUse.Name}'.");

            return null;
        }

        /// <summary>
        /// Retrieves the <see cref="MapView"/> associated with the specified map pane caption,
        /// if the map is currently open in a pane.
        /// </summary>
        /// <param name="mapName">The name of the map pane (caption) to find the view for.</param>
        /// <returns>
        /// A <see cref="MapView"/> instance if the map caption is found in an open pane; otherwise, <c>null</c>.
        /// </returns>
        public MapView GetMapViewFromName(string mapName)
        {
            // Return null if no map name was provided.
            if (mapName == null)
            {
                TraceLog("GetMapViewFromNameAsync error: No map name provided.");
                return null;
            }

            // Access the UI thread to search for a pane showing the specified map.
            // Only UI thread can access FrameworkApplication.Panes.
            MapView mapView = FrameworkApplication.Panes
                .OfType<IMapPane>()
                .FirstOrDefault(p => p.Caption.Equals(mapName, StringComparison.OrdinalIgnoreCase))
                ?.MapView;

            if (mapView == null)
            {
                TraceLog($"GetMapViewFromNameAsync error: No MapView found for with caption '{mapName}'.");
            }

            // Return the found MapView or null if not found.
            return mapView;
        }

        /// <summary>
        /// Gets the <see cref="MapView"/> associated with the specified <see cref="Map"/>.
        /// </summary>
        /// <param name="map">The map to search for in open panes.</param>
        /// <returns>
        /// A task that returns the <see cref="MapView"/> displaying the map,
        /// or <c>null</c> if no such map view is found.
        /// </returns>
        public async Task<MapView> GetMapViewFromMapAsync(Map map)
        {
            if (map == null)
            {
                TraceLog("GetMapViewFromMapAsync error: No map provided.");
                return null;
            }

            MapView mapView = null;

            await QueuedTask.Run(() =>
            {
                // Loop through all panes and find the first one showing the map.
                mapView = FrameworkApplication.Panes
                    .OfType<IMapPane>()
                    .FirstOrDefault(p => p.MapView?.Map == map)
                    ?.MapView;

                if (mapView == null)
                {
                    TraceLog($"GetMapViewFromMapAsync error: No MapView found for map '{map.Name}'.");
                }
            });

            return mapView;
        }

        /// <summary>
        /// Pauses or resumes drawing for the specified map, or the active map if none is provided.
        /// </summary>
        /// <param name="pause">If <c>true</c>, drawing will be paused; otherwise, drawing will be resumed.</param>
        /// <param name="targetMap">
        /// Optional map to control drawing for. If <c>null</c>, the internally tracked active map is used.
        /// </param>
        public void PauseDrawing(bool pause, Map targetMap = null)
        {
            // Use the provided map or fall back to the internally stored active map.
            Map mapToUse = targetMap ?? _activeMap;

            // Attempt to retrieve the MapView for the specified map.
            MapView mapViewToUse = GetMapViewFromName(mapToUse.Name);
            if (mapViewToUse == null)
            {
                // Log if the view could not be found — the map may not be open.
                TraceLog("PauseDrawingAsync error: MapView not found.");
                return;
            }

            // Pause or resume drawing depending on the input parameter.
            // This can be useful when performing batch updates or long-running edits.
            mapViewToUse.DrawingPaused = pause;
        }

        /// <summary>
        /// Creates a new map with the specified name and optionally sets it as the active map.
        /// </summary>
        /// <param name="mapName">The name of the new map to create.</param>
        /// <param name="setActive">If true, the new map will be set as active. Otherwise, the current map remains active.</param>
        /// <returns>The task result containing the name of the created map, or null if creation failed.</returns>
        public async Task<string> CreateMapAsync(string mapName, bool setActive = true)
        {
            // If no map name is supplied.
            if (string.IsNullOrEmpty(mapName))
            {
                TraceLog("CreateMapAsync error: Map name is null or empty.");
                return null;
            }

            // Save the current active pane.
            Pane currentPane = ProApp.Panes.ActivePane;
            Map newMap = null;

            try
            {
                await QueuedTask.Run(() =>
                {
                    // Create a new map without a basemap.
                    newMap = MapFactory.Instance.CreateMap(mapName, basemap: Basemap.None);
                });

                // Create the map pane (this must be awaited as it's async).
                var newPane = await ProApp.Panes.CreateMapPaneAsync(newMap, MapViewingMode.Map);

                if (setActive)
                {
                    _activeMap = newMap;
                }
                else
                {
                    // Return to the previously active pane if available.
                    currentPane?.Activate();
                }

                return newMap.Name;
            }
            catch (Exception ex)
            {
                // Log the exception and return null.
                TraceLog($"CreateMapAsync error: Failed to create map '{mapName}', Exception: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Adds a layer from a URL to the specified map, or the active map if none is provided.
        /// </summary>
        /// <param name="url">The URL of the layer to add.</param>
        /// <param name="index">The index at which to insert the layer (default is 0).</param>
        /// <param name="layerName">An optional custom name for the layer. If not provided, a name is derived from the URL.</param>
        /// <param name="targetMap">The target map to which the layer will be added. Defaults to the active map if null.</param>
        /// <returns>The task result indicating whether the layer was added successfully.</returns>
        public async Task<bool> AddLayerToMapAsync(string url, int index = 0, string layerName = "", Map targetMap = null)
        {
            // If the URL is null or whitespace, return false.
            if (string.IsNullOrWhiteSpace(url))
            {
                TraceLog("AddLayerToMapAsync error: URL is null or empty.");
                return false;
            }

            // Use the provided map, or fall back to the active map if none is given.
            Map mapToUse = targetMap ?? _activeMap;

            try
            {
                await QueuedTask.Run(() =>
                {
                    // Create a URI object from the input URL.
                    Uri uri = new(url);

                    // Use the filename (without extension) as the default layer name if none is provided.
                    string defaultName = System.IO.Path.GetFileNameWithoutExtension(uri.LocalPath);
                    string nameToUse = string.IsNullOrWhiteSpace(layerName) ? defaultName : layerName;

                    // Check whether a layer with the same name already exists in the map.
                    bool layerExists = mapToUse.Layers
                        .Any(l => l.Name.Equals(nameToUse, StringComparison.OrdinalIgnoreCase));

                    // If not found, create and add the layer at the specified index.
                    if (!layerExists)
                        LayerFactory.Instance.CreateLayer(uri, mapToUse, index, nameToUse);
                });

                return true;
            }
            catch (Exception ex)
            {
                // Log and return false if any exception occurs during the process.
                TraceLog($"AddLayerToMapAsync error: Failed to add layer from URL '{url}', Exception: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Adds a standalone table to the specified map, or to the active map if none is provided.
        /// </summary>
        /// <param name="url">The URL or local path of the table to add.</param>
        /// <param name="index">The index at which to insert the table in the standalone table collection (default is 0).</param>
        /// <param name="tableName">An optional custom name for the table. If not provided, a name is derived from the URL.</param>
        /// <param name="targetMap">The map to add the table to. If null, the active map is used.</param>
        /// <returns>True if the table was added successfully; otherwise, false.</returns>
        public async Task<bool> AddTableToMapAsync(string url, int index = 0, string tableName = "", Map targetMap = null)
        {
            // Validate the input URL.
            if (string.IsNullOrWhiteSpace(url))
            {
                TraceLog("AddTableToMapAsync error: URL is null or empty.");
                return false;
            }

            // Use the provided map or fall back to the active map (guaranteed non-null).
            Map mapToUse = targetMap ?? _activeMap;

            try
            {
                await QueuedTask.Run(() =>
                {
                    // Create a URI from the provided URL.
                    Uri uri = new(url);

                    // Use the filename (without extension) as the default layer name if none is provided.
                    string defaultName = System.IO.Path.GetFileNameWithoutExtension(uri.LocalPath);
                    string nameToUse = string.IsNullOrWhiteSpace(tableName) ? defaultName : tableName;

                    // Check if a table with the same name already exists in the map.
                    bool tableExists = mapToUse.StandaloneTables
                        .Any(t => t.Name.Equals(nameToUse, StringComparison.OrdinalIgnoreCase));

                    // If not found, create and add the standalone table at the specified index.
                    if (!tableExists)
                        StandaloneTableFactory.Instance.CreateStandaloneTable(uri, mapToUse, index, nameToUse);
                });

                return true;
            }
            catch (Exception ex)
            {
                // Log the exception and return false.
                TraceLog($"AddTableToMapAsync error: Failed to add table from URL '{url}', Exception: {ex.Message}");
                return false;
            }
        }

        #endregion Map

        #region Zoom

        /// <summary>
        /// Zooms to a feature in the specified layer using a given scale or distance factor.
        /// </summary>
        /// <param name="layerName">The name of the feature layer to zoom to.</param>
        /// <param name="objectID">The object ID of the feature to zoom to.</param>
        /// <param name="factor">Optional. The zoom factor to apply (e.g., 2.0 for twice the extent size).</param>
        /// <param name="mapScaleOrDistance">Optional. The desired map scale or distance in map units.</param>
        /// <param name="targetMap">Optional. The target map to use. Defaults to the active map if null.</param>
        /// <returns>True if zoom was successful; otherwise, false.</returns>
        public async Task<bool> ZoomToFeatureInMapAsync(
            string layerName,
            long objectID,
            double? factor,
            double? mapScaleOrDistance,
            Map targetMap = null)
        {
            // Check there is an input feature layer name.
            if (string.IsNullOrEmpty(layerName))
            {
                TraceLog("ZoomToFeatureInMapAsync error: Layer name is null or empty.");
                return false;
            }

            // Check if the input factor is valid.
            if (factor.HasValue && factor.Value <= 0)
            {
                TraceLog("ZoomToFeatureInMapAsync error: Factor must be greater than zero.");
                return false;
            }

            // Check if the input mapScaleOrDistance is valid.
            if (mapScaleOrDistance.HasValue && factor.Value <= 0)
            {
                TraceLog("ZoomToFeatureInMapAsync error: MapScaleOrDistance must be greater than zero.");
                return false;
            }

            // Use provided map or default to _activeMap.
            Map mapToUse = targetMap ?? _activeMap;

            // Get the map view associated with the map.
            MapView mapViewToUse = GetMapViewFromName(mapToUse.Name);
            if (mapViewToUse == null)
            {
                TraceLog($"ZoomToFeatureInMapAsync error: Map view for map '{mapToUse.Name}' could not be found.");
                return false;
            }

            // Find the target feature layer.
            var targetLayer = await FindLayerAsync(layerName, mapToUse);
            if (targetLayer is not FeatureLayer featureLayer)
            {
                TraceLog("ZoomToFeatureInMapAsync error: target layer is not a FeatureLayer.");
                return false;
            }

            try
            {
                // Zoom to the extent of the specified object ID.
                await mapViewToUse.ZoomToAsync(
                    featureLayer,
                    objectID,
                    duration: null,
                    maintainViewDirection: true,
                    factor: factor,
                    mapScaleOrDistance: mapScaleOrDistance);
            }
            catch (Exception ex)
            {
                // Log the exception and return false.
                TraceLog($"ZoomToFeatureInMapAsync error: Exception {ex.Message}");
                return false;
            }

            return true;
        }

        /// <summary>
        /// Zooms to the extent of specified object IDs in a feature layer.
        /// </summary>
        /// <param name="layerName">The name of the layer containing the objects.</param>
        /// <param name="objectIDs">A list of object IDs to zoom to.</param>
        /// <param name="targetMap">Optional target map; defaults to _activeMap.</param>
        /// <returns>True if zoom succeeded; false otherwise.</returns>
        public async Task<bool> ZoomToFeaturesInMapAsync(string layerName,
            IEnumerable<long> objectIDs,
            Map targetMap = null)
        {
            // Check there is an input feature layer name.
            if (string.IsNullOrEmpty(layerName))
            {
                TraceLog("ZoomToFeaturesInMapAsync error: No layer name provided.");
                return false;
            }

            // Check if there are any input objects.
            if (objectIDs == null || !objectIDs.Any())
            {
                TraceLog("ZoomToFeaturesInMapAsync error: No object IDs provided.");
                return false;
            }

            // Use provided map or default to _activeMap.
            Map mapToUse = targetMap ?? _activeMap;

            // Get the map view associated with the map.
            MapView mapViewToUse = GetMapViewFromName(mapToUse.Name);
            if (mapViewToUse == null)
            {
                TraceLog($"ZoomToFeaturesInMapAsync error: Map view '{mapToUse.Name}' could not be found.");
                return false;
            }

            // Find the target feature layer.
            var targetLayer = await FindLayerAsync(layerName, mapToUse);
            if (targetLayer is not FeatureLayer featureLayer)
            {
                TraceLog($"ZoomToFeaturesInMapAsync error: Feature layer '{layerName}' not found in map.");
                return false;
            }

            try
            {
                // Zoom to the extent of the specified object IDs.
                await mapViewToUse.ZoomToAsync(featureLayer,
                    objectIDs,
                    duration: null,
                    maintainViewDirection: true);
            }
            catch (Exception ex)
            {
                // Log the exception and return false.
                TraceLog($"ZoomToFeaturesInMapAsync error: Exception {ex.Message}");
                return false;
            }

            return true;
        }

        /// <summary>
        /// Zooms to the extent of a layer in a map view for a given ratio or nearest valid scale.
        /// </summary>
        /// <param name="layerName">The name of the layer to zoom to.</param>
        /// <param name="selectedOnly">If true, zooms to selected features only.</param>
        /// <param name="ratio">Optional zoom ratio multiplier.</param>
        /// <param name="scale">Optional fixed scale to set after zooming.</param>
        /// <param name="targetMap">Optional map to use; defaults to _activeMap.</param>
        /// <returns>True if zoom succeeded; false otherwise.</returns>
        public async Task<bool> ZoomToLayerInMapAsync(string layerName,
            bool selectedOnly,
            double? ratio = 1,
            double? scale = 10000,
            Map targetMap = null)
        {
            if (string.IsNullOrEmpty(layerName))
            {
                TraceLog("ZoomToLayerInMapAsync error: No layer name provided.");
                return false;
            }

            if (ratio.HasValue && ratio.Value <= 0)
            {
                TraceLog($"ZoomToLayerInMapAsync error: Invalid zoom ratio: {ratio}.");
                return false;
            }

            if (scale.HasValue && scale.Value <= 0)
            {
                TraceLog($"ZoomToLayerInMapAsync error: Invalid zoom scale: {scale}.");
                return false;
            }

            Map mapToUse = targetMap ?? _activeMap;
            MapView mapViewToUse = GetMapViewFromName(mapToUse.Name);
            if (mapViewToUse == null)
            {
                TraceLog($"ZoomToLayerInMapAsync error: Map view '{mapToUse.Name}' could not be found.");
                return false;
            }

            Layer targetLayer = await FindLayerAsync(layerName, mapToUse);
            if (targetLayer == null)
            {
                TraceLog($"ZoomToLayerInMapAsync error: Layer '{layerName}' not found in map.");
                return false;
            }

            try
            {
                // Zoom to the extent of the layer or its selection.
                await mapViewToUse.ZoomToAsync(targetLayer, selectedOnly);

                // Get the current camera.
                var camera = mapViewToUse.Camera;

                // Apply ratio or fixed scale (mutually exclusive).
                if (ratio.HasValue)
                {
                    camera.Scale *= (double)ratio;
                }
                else if (scale.HasValue)
                {
                    camera.Scale = (double)scale;
                }

                // Apply the modified camera.
                await mapViewToUse.ZoomToAsync(camera, duration: null);
            }
            catch (Exception ex)
            {
                // Log the exception and return false.
                TraceLog($"ZoomToLayerInMapAsync error: Exception {ex.Message}");
                return false;
            }

            return true;
        }

        #endregion Zoom

        //TODO: Finish improving the code and add more comments.

        #region Layers

        /// <summary>
        /// Find a feature layer by name in the active map.
        /// </summary>
        /// <param name="layerName">The name of the layer to find.</param>
        /// <param name="targetMap">The map to search; if null, the active map is used.</param>
        /// <returns>FeatureLayer</returns>
        internal async Task<FeatureLayer> FindLayerAsync(string layerName, Map targetMap = null)
        {
            // Check there is an input feature layer name.
            if (string.IsNullOrEmpty(layerName))
            {
                TraceLog("FindLayer error: No layer name provided.");
                return null;
            }

            // Use provided map or default to _activeMap.
            Map mapToUse = targetMap ?? _activeMap;

            try
            {
                return await QueuedTask.Run(() =>
                {
                    return mapToUse.FindLayers(layerName, true)
                                   .OfType<FeatureLayer>()
                                   .FirstOrDefault();
                });
            }
            catch (Exception ex)
            {
                // Log the exception and return null.
                TraceLog($"FindLayer error: Exception {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Find the position index for a feature layer by name in the active map.
        /// </summary>
        /// <param name="layerName">The name of the layer to find.</param>
        /// <param name="targetMap">The map to search; if null, the active map is used.</param>
        /// <returns>The index of the layer, or 0 if not found.</returns>
        internal async Task<int> FindLayerIndexAsync(string layerName, Map targetMap = null)
        {
            // Check there is an input feature layer name.
            if (string.IsNullOrEmpty(layerName))
            {
                TraceLog("FindLayerIndexAsync error: No layer name provided.");
                return 0;
            }

            // Use provided map or default to _activeMap.
            Map mapToUse = targetMap ?? _activeMap;

            try
            {
                // Run on the CIM thread to safely access layer properties and collection.
                return await QueuedTask.Run(() =>
                {
                    // Iterate through all layers in the map.
                    for (int index = 0; index < mapToUse.Layers.Count; index++)
                    {
                        // Get the index of the first feature layer found by name.
                        // Access to Layer.Name must occur on the CIM thread.
                        if (mapToUse.Layers[index].Name == layerName)
                            return index;
                    }

                    // If no layer matched, return 0 as the default.
                    {
                        TraceLog("FindLayerIndexAsync error: No matching layer found.");
                        return 0;
                    }
                });
            }
            catch (Exception ex)
            {
                // Log the exception and return 0.
                TraceLog($"FindLayerIndexAsync error: Exception {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// Remove a layer by name from the active map.
        /// </summary>
        /// <param name="layerName">The name of the layer to remove.</param>
        /// <param name="targetMap">The map to remove the layer from; if null, the active map is used.</param>
        /// <returns>True if removal succeeded; false otherwise.</returns>
        public async Task<bool> RemoveLayerAsync(string layerName, Map targetMap = null)
        {
            // Check there is an input layer name.
            if (string.IsNullOrEmpty(layerName))
            {
                TraceLog("RemoveLayerAsync error: No layer name provided.");
                return false;
            }

            // Use provided map or default to _activeMap.
            Map mapToUse = targetMap ?? _activeMap;

            try
            {
                // Find the layer in the active map.
                FeatureLayer layer = await FindLayerAsync(layerName, mapToUse);

                // Remove the layer.
                if (layer != null)
                    return await RemoveLayerAsync(layer, mapToUse);
            }
            catch (Exception ex)
            {
                // Log the exception and return false.
                TraceLog($"RemoveLayerAsync error: Exception {ex.Message}");
                return false;
            }

            return true;
        }

        /// <summary>
        /// Remove a layer from the active map.
        /// </summary>
        /// <param name="layer">The layer to remove.</param>
        /// <param name="targetMap">The map to remove the layer from; if null, the active map is used.</param>
        /// <returns>True if removal succeeded; false otherwise.</returns>
        public async Task<bool> RemoveLayerAsync(Layer layer, Map targetMap = null)
        {
            // Check there is an input layer.
            if (layer == null)
            {
                TraceLog("RemoveLayerAsync error: No layer provided.");
                return false;
            }

            // Use provided map or default to _activeMap.
            Map mapToUse = targetMap ?? _activeMap;

            try
            {
                await QueuedTask.Run(() =>
                {
                    // Remove the layer.
                    if (layer != null)
                        mapToUse.RemoveLayer(layer);
                });
            }
            catch (Exception ex)
            {
                // Log the exception and return false.
                TraceLog($"RemoveLayerAsync error: Exception {ex.Message}");
                return false;
            }

            return true;
        }

        #endregion Layers

        #region Select

        /// <summary>
        /// Select features in feature class by location.
        /// </summary>
        /// <param name="targetLayer">The target layer to select features in.</param>
        /// <param name="searchLayer">The search layer to use for spatial selection.</param>
        /// <param name="overlapType">The type of spatial relationship (e.g., INTERSECT, CONTAINS).</param>
        /// <param name="searchDistance">The search distance for buffering (optional).</param>
        /// <param name="selectionType">The selection type (e.g., NEW_SELECTION, ADD_TO_SELECTION).</param>
        /// <returns>True if selection succeeded; false otherwise.</returns>
        public static async Task<bool> SelectLayerByLocationAsync(string targetLayer, string searchLayer,
            string overlapType = "INTERSECT", string searchDistance = "", string selectionType = "NEW_SELECTION")
        {
            // Check if there is an input target layer name.
            if (string.IsNullOrEmpty(targetLayer))
            {
                TraceLog("SelectLayerByLocationAsync error: No target layer name provided.");
                return false;
            }

            // Check if there is an input search layer name.
            if (string.IsNullOrEmpty(searchLayer))
            {
                TraceLog("SelectLayerByLocationAsync error: No search layer name provided.");
                return false;
            }

            // Make a value array of strings to be passed to the tool.
            IReadOnlyList<string> parameters = Geoprocessing.MakeValueArray(targetLayer, overlapType, searchLayer, searchDistance, selectionType);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geoprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; // | GPExecuteToolFlags.RefreshProjectItems;

            //Geoprocessing.OpenToolDialog("management.SelectLayerByLocation", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.SelectLayerByLocation", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception ex)
            {
                // Log the exception and return false.
                TraceLog($"SelectLayerByLocationAsync error: Exception occurred while selecting features. TargetLayer: {targetLayer}, SearchLayer: {searchLayer}, Exception: {ex.Message}");
                return false;
            }

            return true;
        }

        /// <summary>
        /// Select features in feature class by location.
        /// </summary>
        /// <param name="targetLayer">The target layer to select features in.</param>
        /// <param name="searchLayer">The search layer to use for spatial selection.</param>
        /// <param name="overlapType">The type of spatial relationship (e.g., INTERSECT, CONTAINS).</param>
        /// <param name="searchDistance">The search distance for buffering (optional).</param>
        /// <param name="selectionType">The selection type (e.g., NEW_SELECTION, ADD_TO_SELECTION).</param>
        /// <returns>The task result is true if selection succeeded; false otherwise.</returns>
        public static async Task<bool> SelectLayerByLocationAsync(FeatureLayer targetLayer, FeatureLayer searchLayer,
            string overlapType = "INTERSECT", string searchDistance = "", string selectionType = "NEW_SELECTION")
        {
            // Check there is an input feature layer.
            if (targetLayer == null)
            {
                TraceLog("SelectLayerByLocationAsync error: No target layer name provided.");
                return false;
            }

            // Check there is an input search layer.
            if (searchLayer == null)
            {
                TraceLog("SelectLayerByLocationAsync error: No search layer name provided.");
                return false;
            }

            return await QueuedTask.Run(() =>
            {
                // Attempt to get the selected ObjectIDs in the search layer.
                var oidSet = searchLayer.GetSelection()?.GetObjectIDs();

                // Use a query filter — either for selected features or all features.
                QueryFilter queryFilter;

                // If any selected features to build the geometry.
                if (oidSet != null && oidSet.Count > 0)
                {
                    // Use only selected features.
                    queryFilter = new QueryFilter
                    {
                        ObjectIDs = oidSet
                    };
                }
                else
                {
                    // No selected features — fallback to using all features.
                    queryFilter = new QueryFilter();
                }

                // Union geometry of the features in the search layer to use as spatial filter.
                Geometry searchGeometry;

                using (var rowCursor = searchLayer.Search(queryFilter))
                {
                    var geometries = new List<Geometry>();

                    while (rowCursor.MoveNext())
                    {
                        using var feature = rowCursor.Current as Feature;
                        if (feature?.GetShape() != null)
                            geometries.Add(feature.GetShape());
                    }

                    if (geometries.Count == 0)
                    {
                        TraceLog("SelectLayerByLocationAsync error: No geometry found in search layer.");
                        return false;
                    }

                    searchGeometry = GeometryEngine.Instance.Union(geometries);
                }

                if (searchGeometry == null)
                {
                    TraceLog("SelectLayerByLocationAsync error: No geometry found in search layer.");
                    return false;
                }

                // Optionally buffer the search geometry if a distance is provided.
                if (!string.IsNullOrEmpty(searchDistance) && double.TryParse(searchDistance, out double distance) && distance > 0)
                {
                    // Use the spatial reference of the search geometry to maintain units.
                    var spatialRef = searchGeometry.SpatialReference;

                    // Buffer assumes units match geometry’s spatial reference (e.g., meters if projected).
                    searchGeometry = GeometryEngine.Instance.Buffer(searchGeometry, distance);

                    if (searchGeometry == null)
                    {
                        TraceLog("SelectLayerByLocationAsync error: Buffering search geometry failed.");
                        return false;
                    }
                }

                // Map string overlapType to SpatialRelationship.
                SpatialRelationship spatialRel = overlapType.ToUpper() switch
                {
                    "INTERSECT" => SpatialRelationship.Intersects,
                    "CONTAINS" => SpatialRelationship.Contains,
                    "WITHIN" => SpatialRelationship.Within,
                    "CROSSES" => SpatialRelationship.Crosses,
                    "TOUCHES" => SpatialRelationship.Touches,
                    "OVERLAPS" => SpatialRelationship.Overlaps,
                    _ => SpatialRelationship.Intersects
                };

                // Prepare the spatial query.
                var spatialFilter = new SpatialQueryFilter
                {
                    FilterGeometry = searchGeometry,
                    SpatialRelationship = spatialRel
                };

                // Determine selection combination method.
                SelectionCombinationMethod method = selectionType.ToUpper() switch
                {
                    "ADD_TO_SELECTION" => SelectionCombinationMethod.Add,
                    "REMOVE_FROM_SELECTION" => SelectionCombinationMethod.Subtract,
                    "SELECT_NEW" or "NEW_SELECTION" => SelectionCombinationMethod.New,
                    "INTERSECT_WITH_SELECTION" => SelectionCombinationMethod.And,
                    _ => SelectionCombinationMethod.New
                };

                // Perform the selection.
                targetLayer.Select(spatialFilter, method);

                return true;
            });
        }

        /// <summary>
        /// Select features in layerName by attributes.
        /// </summary>
        /// <param name="layerName">The name of the layer to select features in.</param>
        /// <param name="whereClause">The
        /// <param name="selectionMethod">The selection combination method.</param>
        /// <returns>The task result is true if selection succeeded; false otherwise.</returns>
        public async Task<bool> SelectLayerByAttributesAsync(string layerName, string whereClause, SelectionCombinationMethod selectionMethod = SelectionCombinationMethod.New, Map targetMap = null)
        {
            // Check there is an input feature layer name.
            if (string.IsNullOrEmpty(layerName))
            {
                TraceLog("SelectLayerByAttributesAsync error: No input layer name provided.");
                return false;
            }

            try
            {
                // Find the feature layerName by name if it exists. Only search existing layers.
                FeatureLayer featureLayer = await FindLayerAsync(layerName, targetMap);

                if (featureLayer == null)
                {
                    TraceLog("SelectLayerByAttributesAsync error: Feature layer not found.");
                    return false;
                }

                // Create a query filter using the where clause.
                QueryFilter queryFilter = new()
                {
                    WhereClause = whereClause
                };

                await QueuedTask.Run(() =>
                {
                    // Select the features matching the search clause.
                    featureLayer.Select(queryFilter, selectionMethod);
                });
            }
            catch (Exception ex)
            {
                // Log the exception and return false.
                TraceLog($"SelectLayerByAttributesAsync error: Exception occurred while selecting features. Layer: {layerName}, WhereClause: {whereClause}, Exception: {ex.Message}");
                return false;
            }

            return true;
        }

        /// <summary>
        /// Clear selected features in a feature layer.
        /// </summary>
        /// <param name="layerName">The name of the layer to clear selection in.</param>
        /// <returns>The task result is true if clearing selection succeeded; false otherwise.</returns>
        public async Task<bool> ClearLayerSelectionAsync(string layerName, Map targetMap = null)
        {
            // Check there is an input feature layer name.
            if (string.IsNullOrEmpty(layerName))
                return false;

            try
            {
                // Find the feature layerName by name if it exists. Only search existing layers.
                FeatureLayer featureLayer = await FindLayerAsync(layerName, targetMap);

                if (featureLayer == null)
                    return false;

                await QueuedTask.Run(() =>
                {
                    // Clear the feature selection.
                    featureLayer.ClearSelection();
                });
            }
            catch (Exception ex)
            {
                // Log the exception and return false.
                TraceLog($"ClearLayerSelectionAsync error: Exception occurred while clearing selection. Layer: {layerName}, Exception: {ex.Message}");
                return false;
            }

            return true;
        }

        /// <summary>
        /// Count the number of selected features in a feature layer.
        /// </summary>
        /// <param name="layerName">The name of the layer to count selected features in.</param>
        /// <returns>The task result is the count of selected features, or -1 on error.</returns>
        public async Task<long> GetSelectedFeatureCountAsync(string layerName, Map targetMap = null)
        {
            // Check there is an input feature layer name.
            if (string.IsNullOrEmpty(layerName))
                return -1;

            long selectedCount;
            try
            {
                // Find the feature layerName by name if it exists. Only search existing layers.
                FeatureLayer featureLayer = await FindLayerAsync(layerName, targetMap);

                if (featureLayer == null)
                    return -1;

                // Select the features matching the search clause.
                selectedCount = await QueuedTask.Run(() => featureLayer.SelectionCount);
            }
            catch (Exception ex)
            {
                // Log the exception and return -1.
                TraceLog($"GetSelectedFeatureCount error: Exception occurred while counting selected features. Layer: {layerName}, Exception: {ex.Message}");
                return -1;
            }

            return selectedCount;
        }

        /// <summary>
        /// Get the list of fields for a feature class.
        /// </summary>
        /// <param name="layerPath"></param>
        /// <returns>IReadOnlyList<Field></returns>
        public async Task<IReadOnlyList<Field>> GetFCFieldsAsync(string layerPath, Map targetMap = null)
        {
            // Check there is an input feature layer path.
            if (string.IsNullOrEmpty(layerPath))
                return null;

            try
            {
                // Find the feature layer by name if it exists. Only search existing layers.
                FeatureLayer featureLayer = await FindLayerAsync(layerPath, targetMap);

                if (featureLayer == null)
                    return null;

                IReadOnlyList<Field> fields = null;
                List<string> fieldList = [];

                await QueuedTask.Run(() =>
                {
                    // Get the underlying feature class as a table.
                    using Table table = featureLayer.GetTable();
                    if (table != null)
                    {
                        // Get the table definition of the table.
                        using TableDefinition tableDef = table.GetDefinition();

                        // Get the fields in the table.
                        fields = tableDef.GetFields();
                    }
                });

                return fields;
            }
            catch (Exception ex)
            {
                // Log the exception and return null.
                TraceLog($"GetFCFieldsAsync error: Exception occurred while getting fields. Layer: {layerPath}, Exception: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Get the list of fields for a standalone table.
        /// </summary>
        /// <param name="layerPath">The standalone table name.</param>
        /// <param name="targetMap">The map to search; if null, the active map is used.</param>
        /// <returns>The task result is IReadOnlyList<ArcGIS.Core.Data.Field></returns>
        public async Task<IReadOnlyList<Field>> GetTableFieldsAsync(string layerPath, Map targetMap = null)
        {
            // Check there is an input feature layer name.
            if (string.IsNullOrEmpty(layerPath))
                return null;

            try
            {
                // Find the table by name if it exists. Only search existing layers.
                StandaloneTable inputTable = FindTable(layerPath, targetMap);

                if (inputTable == null)
                    return null;

                IReadOnlyList<Field> fields = null;
                List<string> fieldList = [];

                await QueuedTask.Run(() =>
                {
                    // Get the underlying table.
                    using Table table = inputTable.GetTable();
                    if (table != null)
                    {
                        // Get the table definition of the table.
                        using TableDefinition tableDef = table.GetDefinition();

                        // Get the fields in the table.
                        fields = tableDef.GetFields();
                    }
                });

                return fields;
            }
            catch (Exception ex)
            {
                // Log the exception and return null.
                TraceLog($"GetTableFieldsAsync error: Exception occurred while getting fields. Layer: {layerPath}, Exception {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Determines whether all specified field names exist in the provided collection of fields.
        /// </summary>
        /// <remarks>If <paramref name="fieldNames"/> is null or empty, the method returns false. Field
        /// names are trimmed of whitespace before checking for existence.</remarks>
        /// <param name="fields">The collection of fields to search.</param>
        /// <param name="fieldNames">A delimited string containing the names of the fields to check for existence. Field names are separated by
        /// the specified separator character or a comma.</param>
        /// <param name="separator">The character used to separate field names in the <paramref name="fieldNames"/> string.</param>
        /// <param name="checkStrings">Indicates whether to perform additional string-based checks when verifying field existence. The default is
        /// <see langword="false"/>.</param>
        /// <returns>True if all specified field names exist in the collection; otherwise, false.</returns>
        public static bool FieldsExist(IReadOnlyList<Field> fields, string fieldNames, char separator, bool checkStrings = false)
        {
            // Check there is an input field name.
            if (string.IsNullOrEmpty(fieldNames))
                return false;

            // Split the field names into a list.
            string[] fieldNameArray = fieldNames.Split([',', separator], StringSplitOptions.RemoveEmptyEntries);
            foreach (string fieldName in fieldNameArray)
            {
                if (!FieldExists(fields, fieldName.Trim()))
                    return false;
            }
            return true;
        }

        /// <summary>
        /// Check if a field exists in a list of fields.
        /// </summary>
        /// <param name="fields">The list of fields to check.</param>
        /// <param name="fieldName">The name of the field to check for existence.</param>
        /// <returns>True if field exists; false otherwise.</returns>
        public static bool FieldExists(IReadOnlyList<Field> fields, string fieldName)
        {
            bool fldFound = false;

            // Check there is an input field name.
            if (string.IsNullOrEmpty(fieldName))
                return false;

            foreach (Field fld in fields)
            {
                if (fld.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase) ||
                    (fld.AliasName != null && fld.AliasName.Equals(fieldName, StringComparison.OrdinalIgnoreCase)))
                {
                    fldFound = true;
                    break;
                }
            }

            return fldFound;
        }

        /// <summary>
        /// Check if a field exists in a feature class.
        /// </summary>
        /// <param name="layerPath">The feature layer path.</param>
        /// <param name="fieldName">The name of the field to check for existence.</param>
        /// <returns>The task result is true if field exists; false otherwise.</returns>
        public async Task<bool> FieldExistsAsync(string layerPath, string fieldName, Map targetMap = null)
        {
            // Check there is an input feature layer path.
            if (string.IsNullOrEmpty(layerPath))
                return false;

            // Check there is an input field name.
            if (string.IsNullOrEmpty(fieldName))
                return false;

            try
            {
                // Find the feature layer by name if it exists. Only search existing layers.
                FeatureLayer featureLayer = await FindLayerAsync(layerPath, targetMap);

                if (featureLayer == null)
                    return false;

                bool fldFound = false;

                await QueuedTask.Run(() =>
                {
                    // Get the underlying feature class as a table.
                    using Table table = featureLayer.GetTable();
                    if (table != null)
                    {
                        // Get the table definition of the table.
                        using TableDefinition tableDef = table.GetDefinition();

                        // Get the fields in the table.
                        IReadOnlyList<Field> fields = tableDef.GetFields();

                        // Loop through all fields looking for a name match.
                        foreach (Field fld in fields)
                        {
                            if (fld.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase) ||
                                (fld.AliasName != null && fld.AliasName.Equals(fieldName, StringComparison.OrdinalIgnoreCase)))
                            {
                                fldFound = true;
                                break;
                            }
                        }
                    }
                });

                return fldFound;
            }
            catch (Exception ex)
            {
                // Log the exception and return false.
                TraceLog($"FieldExistsAsync error: Exception occurred while checking field existence. Layer: {layerPath}, Field: {fieldName}, Exception {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Check if a list of fields exists in a feature class and
        /// return a list of those that do.
        /// </summary>
        /// <param name="layerName">The feature layer name.</param>
        /// <param name="fieldNames">The list of field names to check for existence.</param>
        /// <returns>The task result is List<string></returns>
        public async Task<List<string>> GetExistingFieldsAsync(string layerName, List<string> fieldNames, Map targetMap = null)
        {
            List<string> fieldsThatExist = [];
            foreach (string fieldName in fieldNames)
            {
                if (await FieldExistsAsync(layerName, fieldName, targetMap))
                    fieldsThatExist.Add(fieldName);
            }

            return fieldsThatExist;
        }

        /// <summary>
        /// Check if a field is numeric in a feature class.
        /// </summary>
        /// <param name="layerName">The feature layer name.</param>
        /// <param name="fieldName">The name of the field to check.</param>
        /// <returns>The task result is true if field is numeric; false otherwise.</returns>
        public async Task<bool> FieldIsNumericAsync(string layerName, string fieldName, Map targetMap = null)
        {
            // Check there is an input feature layer name.
            if (string.IsNullOrEmpty(layerName))
                return false;

            // Check there is an input field name.
            if (string.IsNullOrEmpty(fieldName))
                return false;

            try
            {
                // Find the feature layerName by name if it exists. Only search existing layers.
                FeatureLayer featureLayer = await FindLayerAsync(layerName, targetMap);

                if (featureLayer == null)
                    return false;

                IReadOnlyList<Field> fields = null;

                bool fldIsNumeric = false;

                await QueuedTask.Run(() =>
                {
                    // Get the underlying feature class as a table.
                    using Table table = featureLayer.GetTable();
                    if (table != null)
                    {
                        // Get the table definition of the table.
                        using TableDefinition tableDef = table.GetDefinition();

                        // Get the fields in the table.
                        fields = tableDef.GetFields();

                        // Loop through all fields looking for a name match.
                        foreach (Field fld in fields)
                        {
                            if (fld.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase) ||
                                (fld.AliasName != null && fld.AliasName.Equals(fieldName, StringComparison.OrdinalIgnoreCase)))
                            {
                                fldIsNumeric = fld.FieldType switch
                                {
                                    FieldType.SmallInteger => true,
                                    FieldType.BigInteger => true,
                                    FieldType.Integer => true,
                                    FieldType.Single => true,
                                    FieldType.Double => true,
                                    _ => false,
                                };

                                break;
                            }
                        }
                    }
                });

                return fldIsNumeric;
            }
            catch (Exception ex)
            {
                // Log the exception and return false.
                TraceLog($"FieldIsNumericAsync error: Exception occurred while checking field type. Layer: {layerName}, Field: {fieldName}, Exception {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Calculate the total row length for a feature class
        /// </summary>
        /// <param name="layerName">The feature layer name.</param>
        /// <returns>The task result is the row length as an integer.</returns>
        public async Task<int> GetFCRowLengthAsync(string layerName, Map targetMap = null)
        {
            // Check there is an input feature layer name.
            if (string.IsNullOrEmpty(layerName))
                return 0;

            try
            {
                // Find the feature layerName by name if it exists. Only search existing layers.
                FeatureLayer featureLayer = await FindLayerAsync(layerName, targetMap);

                if (featureLayer == null)
                    return 0;

                IReadOnlyList<Field> fields = null;
                List<string> fieldList = [];

                int rowLength = 1;

                await QueuedTask.Run(() =>
                {
                    // Get the underlying feature class as a table.
                    using Table table = featureLayer.GetTable();
                    if (table != null)
                    {
                        // Get the table definition of the table.
                        using TableDefinition tableDef = table.GetDefinition();

                        // Get the fields in the table.
                        fields = tableDef.GetFields();

                        int fldLength;

                        // Loop through all fields.
                        foreach (Field fld in fields)
                        {
                            if (fld.FieldType == FieldType.Integer)
                                fldLength = 10;
                            else if (fld.FieldType == FieldType.Geometry)
                                fldLength = 0;
                            else
                                fldLength = fld.Length;

                            rowLength += fldLength;
                        }
                    }
                });

                return rowLength;
            }
            catch (Exception ex)
            {
                // Log the exception and return 0.
                TraceLog($"GetFCRowLengthAsync error: Exception occurred while getting row length. Layer: {layerName}, Exception {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// Deletes all the fields from a feature class that are not required.
        /// </summary>
        /// <param name="layerName">The feature layer name.</param>
        /// <param name="fieldList">The list of field names to keep.</param>
        /// <returns>The task result is true if fields were deleted; false otherwise.</returns>
        public async Task<bool> KeepSelectedFieldsAsync(string layerName, List<string> fieldList, Map targetMap = null)
        {
            // Check the input parameters.
            if (string.IsNullOrEmpty(layerName))
                return false;

            if (fieldList == null || fieldList.Count == 0)
                return false;

            // Add a FID field so that it isn't tried to be removed.
            //fieldList.Add("FID");

            // Get the list of fields for the input table.
            IReadOnlyList<Field> inputfields = await GetFCFieldsAsync(layerName, targetMap);

            // Check a list of fields is returned.
            if (inputfields == null || inputfields.Count == 0)
                return false;

            // Get the list of field names for the input table that
            // aren't required fields (e.g. excluding FID and Shape).
            List<string> inputFieldNames = [.. inputfields.Where(x => !x.IsRequired).Select(y => y.Name)];

            // Get the list of fields that do exist in the layer.
            List<string> existingFields = await GetExistingFieldsAsync(layerName, fieldList, targetMap);

            // Get the list of layer fields that aren't in the field list.
            var remainingFields = inputFieldNames.Except(existingFields).ToList();

            if (remainingFields == null || remainingFields.Count == 0)
                return true;

            // Make a value array of strings to be passed to the tool.
            var parameters = Geoprocessing.MakeValueArray(layerName, remainingFields);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geoprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; //| GPExecuteToolFlags.RefreshProjectItems;

            //Geoprocessing.OpenToolDialog("management.DeleteField", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.DeleteField", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception ex)
            {
                // Log the exception and return false.
                TraceLog($"KeepSelectedFieldsAsync error: Exception occurred while deleting fields. Layer: {layerName}, Exception {ex.Message}");
                return false;
            }

            return true;
        }

        /// <summary>
        /// Get the full layer path name for a layer in the map (i.e.
        /// to include any parent group names).
        /// </summary>
        /// <param name="layer">The layer to get the path for.</param>
        /// <returns>The task result is the full layer path as a string.</returns>
        public Task<string> GetLayerPathAsync(Layer layer)
        {
            return QueuedTask.Run(async () =>
            {
                // Check there is an input layer.
                if (layer == null)
                    return null;

                string layerPath = "";

                try
                {
                    // Get the parent for the layer.
                    ILayerContainer layerParent = layer.Parent;

                    // Loop while the parent is a group layer.
                    while (layerParent is GroupLayer)
                    {
                        // Get the parent layer.
                        Layer groupLayer = (Layer)layerParent;

                        // Append the parent name to the full layer path.
                        // Access to groupLayer.Name must occur on the CIM thread.
                        layerPath = groupLayer.Name + "/" + layerPath;

                        // Get the parent for the layer.
                        layerParent = groupLayer.Parent;
                    }

                    // Append the layer name to its full path.
                    // Access to Layer.Name must occur on the CIM thread.
                    layerPath += layer.Name;
                }
                catch (Exception ex)
                {
                    // Access to Layer.Name must occur on the CIM thread.
                    string safeLayerName = await QueuedTask.Run(() => layer.Name);

                    // Log the exception and return false.
                    TraceLog($"GetLayerPathAsync error: Exception occurred while getting layer path. Layer: {safeLayerName}, Exception: {ex.Message}");
                    return null;
                }

                return layerPath;
            });
        }

        /// <summary>
        /// Get the full layer path name for a layer name in the map (i.e.
        /// to include any parent group names.
        /// </summary>
        /// <param name="layerName">The name of the layer to get the path for.</param>
        /// <returns>The task result is the full layer path as a string.</returns>
        public async Task<string> GetLayerPathAsync(string layerName, Map targetMap = null)
        {
            // Check there is an input layer name.
            if (string.IsNullOrEmpty(layerName))
                return null;

            // Use provided map or default to _activeMap.
            Map mapToUse = targetMap ?? _activeMap;

            try
            {
                FeatureLayer layer = await FindLayerAsync(layerName, mapToUse);
                if (layer == null)
                    return null;

                return await GetLayerPathAsync(layer);
            }
            catch (Exception ex)
            {
                // Log the exception and return false.
                TraceLog($"GetLayerPathAsync error: Exception occurred while getting layer path. Layer: {layerName}, Exception {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Returns a simplified feature class shape type for a feature layer.
        /// </summary>
        /// <param name="featureLayer">The feature layer.</param>
        /// <returns>The task result is string: point, line, polygon</returns>
        public async Task<string> GetFeatureClassTypeAsync(FeatureLayer featureLayer)
        {
            // Check there is an input feature layer.
            if (featureLayer == null)
                return null;

            try
            {
                esriGeometryType shapeType = await QueuedTask.Run(() => featureLayer.ShapeType);

                return shapeType switch
                {
                    esriGeometryType.esriGeometryPoint => "point",
                    esriGeometryType.esriGeometryMultipoint => "point",
                    esriGeometryType.esriGeometryPolygon => "polygon",
                    esriGeometryType.esriGeometryRing => "polygon",
                    esriGeometryType.esriGeometryLine => "line",
                    esriGeometryType.esriGeometryPolyline => "line",
                    esriGeometryType.esriGeometryCircularArc => "line",
                    esriGeometryType.esriGeometryEllipticArc => "line",
                    esriGeometryType.esriGeometryBezier3Curve => "line",
                    esriGeometryType.esriGeometryPath => "line",
                    _ => "other",
                };
            }
            catch (Exception ex)
            {
                // Access to Layer.Name must occur on the CIM thread.
                string safeLayerName = await QueuedTask.Run(() => featureLayer.Name);

                // Log the exception and return null.
                TraceLog($"GetFeatureClassTypeAsync error: Exception occurred while getting shape type. Layer: {safeLayerName}, Exception: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Returns a simplified feature class shape type for a layer name.
        /// </summary>
        /// <param name="layerName">The feature layer name.</param>
        /// <returns>The task result is string: point, line, polygon</returns>
        public async Task<string> GetFeatureClassTypeAsync(string layerName, Map targetMap = null)
        {
            // Check there is an input feature layer name.
            if (string.IsNullOrEmpty(layerName))
                return null;

            // Use provided map or default to _activeMap.
            Map mapToUse = targetMap ?? _activeMap;

            try
            {
                // Find the layer in the active map.
                FeatureLayer layer = await FindLayerAsync(layerName, mapToUse);

                if (layer == null)
                    return null;

                return await GetFeatureClassTypeAsync(layer);
            }
            catch (Exception ex)
            {
                // Log the exception and return null.
                TraceLog($"GetFeatureClassType error: Exception occurred while getting feature class type. Layer: {layerName}, Exception: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Recursively retrieves all feature layers from a collection of layers,
        /// including those nested within group layers.
        /// </summary>
        /// <param name="layers">The layer collection to search.</param>
        /// <returns>All FeatureLayer instances found within the collection.</returns>
        private static IEnumerable<FeatureLayer> GetAllFeatureLayers(IEnumerable<Layer> layers)
        {
            // Loop through each layer in the collection.
            foreach (var layer in layers)
            {
                // If it's a FeatureLayer, return it.
                if (layer is FeatureLayer fl)
                {
                    yield return fl;
                }
                // If it's a GroupLayer, search its children recursively.
                else if (layer is GroupLayer gl)
                {
                    foreach (var child in GetAllFeatureLayers(gl.Layers))
                        yield return child;
                }
            }
        }

        #endregion Layers

        #region Partners

        /// <summary>
        /// Get a list of the active partners from the input layer using the
        /// specified where clause and column names.
        /// </summary>
        /// <param name="inputLayer"></param>
        /// <param name="partnerClause"></param>
        /// <param name="partnerColumn"></param>
        /// <param name="shortColumn"></param>
        /// <param name="notesColumn"></param>
        /// <param name="formatColumn"></param>
        /// <param name="exportColumn"></param>
        /// <param name="sqlTableColumn"></param>
        /// <param name="sqlFilesColumn"></param>
        /// <param name="mapFilesColumn"></param>
        /// <param name="tagsColumn"></param>
        /// <param name="activeColumn"></param>
        /// <returns></returns>
        public async Task<List<Partner>> GetActiveParnersAsync(string inputLayer, string partnerClause, string partnerColumn, string shortColumn, string notesColumn,
            string formatColumn, string exportColumn, string sqlTableColumn, string sqlFilesColumn, string mapFilesColumn, string tagsColumn, string activeColumn)
        {
            // Check there is an input layer name.
            if (string.IsNullOrEmpty(inputLayer))
                return null;

            FeatureLayer inputFeaturelayer;
            List<Partner> partnerList = [];

            try
            {
                // Get the input feature layer.
                inputFeaturelayer = await FindLayerAsync(inputLayer);
                if (inputFeaturelayer == null)
                    return null;

                await QueuedTask.Run(() =>
                {
                    /// Get the feature class for the input feature layer.
                    FeatureClass featureClass = inputFeaturelayer.GetFeatureClass();

                    // Get the feature class defintion.
                    using FeatureClassDefinition featureClassDefinition = featureClass.GetDefinition();

                    // Create a new list of sort descriptions.
                    List<SortDescription> sortDescriptions = [];

                    // Create a query filter using the partner clause.
                    QueryFilter queryFilter = new()
                    {
                        WhereClause = partnerClause,
                        PostfixClause = "ORDER BY " + partnerColumn
                    };

                    // Create a cursor of the sorted features.
                    using RowCursor rowCursor = featureClass.Search(queryFilter, false);

                    // Loop through the feature class/table using the cursor.
                    while (rowCursor.MoveNext())
                    {
                        // Get the current row.
                        using Row record = rowCursor.Current;

                        if (Convert.ToString(record[activeColumn]).ToLower(System.Globalization.CultureInfo.CurrentCulture) is "y")
                        {
                            // Create a new partner for this row.
                            Partner partner = new()
                            {
                                PartnerName = Convert.ToString(record[partnerColumn]),
                                ShortName = Convert.ToString(record[shortColumn]),
                                Notes = Convert.ToString(record[notesColumn]),
                                GISFormat = Convert.ToString(record[formatColumn]),
                                ExportFormat = Convert.ToString(record[exportColumn]),
                                SQLTable = Convert.ToString(record[sqlTableColumn]),
                                SQLFiles = Convert.ToString(record[sqlFilesColumn]),
                                MapFiles = Convert.ToString(record[mapFilesColumn]),
                                Tags = Convert.ToString(record[tagsColumn])
                            };

                            // Add the partner to the list of active partners.
                            partnerList.Add(partner);
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                // Log the exception and return null.
                TraceLog($"GetActiveParnersAsync error: Exception occurred while getting active partners. Layer: {inputLayer}, Exception: {ex.Message}");
                return null;
            }
            finally
            {
            }

            return partnerList;
        }

        #endregion Partners

        #region Group Layers

        /// <summary>
        /// Finds a group layer by name in the specified or active map.
        /// </summary>
        /// <param name="layerName">The name of the group layer to find.</param>
        /// <param name="targetMap">Optional map to search in; defaults to the active map.</param>
        /// <returns>The task result is the GroupLayer if found; otherwise, null.</returns>
        internal async Task<GroupLayer> FindGroupLayerAsync(string layerName, Map targetMap = null)
        {
            // Check there is an input group layer name.
            if (string.IsNullOrEmpty(layerName))
            {
                TraceLog("FindGroupLayerAsync error: No layer name provided.");
                return null;
            }

            // Use provided map or default to _activeMap.
            Map mapToUse = targetMap ?? _activeMap;

            try
            {
                // Run layer lookup on the QueuedTask to comply with ArcGIS Pro threading model.
                return await QueuedTask.Run(() =>
                {
                    return mapToUse.FindLayers(layerName, true)
                                   .OfType<GroupLayer>()
                                   .FirstOrDefault();
                });
            }
            catch (Exception ex)
            {
                // Log the exception and return null.
                TraceLog($"FindGroupLayerAsync error: Exception occurred while finding group layer. Layer: {layerName}, Exception: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Move a layer into a group layer (creating the group layer if
        /// it doesn't already exist).
        /// </summary>
        /// <param name="layer">The layer to move.</param>
        /// <param name="groupLayerName">The name of the group layer to move the layer into.</param>
        /// <param name="position">The position in the group layer to insert the layer; -1 to add to the end.</param>
        /// <returns>The task result is true if successful; otherwise false.</returns>
        public async Task<bool> MoveToGroupLayerAsync(Layer layer, string groupLayerName, int position = -1, Map targetMap = null)
        {
            // Check if there is an input layer.
            if (layer == null)
                return false;

            // Check there is an input group layer name.
            if (string.IsNullOrEmpty(groupLayerName))
                return false;

            // Use provided map or default to _activeMap.
            Map mapToUse = targetMap ?? _activeMap;

            // Does the group layer exist?
            GroupLayer groupLayer = await FindGroupLayerAsync(groupLayerName, mapToUse);
            if (groupLayer == null)
            {
                // Add the group layer to the map.
                try
                {
                    await QueuedTask.Run(() =>
                    {
                        groupLayer = LayerFactory.Instance.CreateGroupLayer(mapToUse, 0, groupLayerName);
                    });
                }
                catch (Exception ex)
                {
                    // Access to Layer.Name must occur on the CIM thread.
                    string safeLayerName = await QueuedTask.Run(() => layer.Name);

                    // Log the exception and return false.
                    TraceLog($"MoveToGroupLayerAsync error: Exception occurred while creating group layer. Layer: {safeLayerName}, GroupLayer: {groupLayerName}, Exception: {ex.Message}");
                    return false;
                }
            }

            // Move the layer into the group.
            try
            {
                await QueuedTask.Run(() =>
                {
                    // Move the layer into the group.
                    mapToUse.MoveLayer(layer, groupLayer, position);

                    // Expand the group.
                    groupLayer.SetExpanded(true);
                });
            }
            catch (Exception ex)
            {
                // Access to Layer.Name must occur on the CIM thread.
                string safeLayerName = await QueuedTask.Run(() => layer.Name);

                // Log the exception and return false.
                TraceLog($"MoveToGroupLayerAsync error: Exception occurred while moving layer to group layer. Layer: {safeLayerName}, GroupLayer: {groupLayerName}, Exception {ex.Message}");
                return false;
            }

            return true;
        }

        /// <summary>
        /// Remove a group layer if it is empty.
        /// </summary>
        /// <param name="groupLayerName">The name of the group layer to remove.</param>
        /// <returns>The task result is true if successful; otherwise false.</returns>
        public async Task<bool> RemoveGroupLayerAsync(string groupLayerName, Map targetMap = null)
        {
            // Check there is an input group layer name.
            if (string.IsNullOrEmpty(groupLayerName))
                return false;

            // Use provided map or default to _activeMap.
            Map mapToUse = targetMap ?? _activeMap;

            try
            {
                // Does the group layer exist?
                GroupLayer groupLayer = await FindGroupLayerAsync(groupLayerName, mapToUse);
                if (groupLayer == null)
                    return false;

                // Count the layers in the group.
                if (groupLayer.Layers.Count != 0)
                    return true;

                await QueuedTask.Run(() =>
                {
                    // Remove the group layer.
                    mapToUse.RemoveLayer(groupLayer);
                });
            }
            catch (Exception ex)
            {
                // Log the exception and return false.
                TraceLog($"RemoveGroupLayerAsync error: Exception occurred while removing group layer. GroupLayer: {groupLayerName}, Exception {ex.Message}");
                return false;
            }

            return true;
        }

        #endregion Group Layers

        #region Tables

        /// <summary>
        /// Find a table by name in the active map.
        /// </summary>
        /// <param name="tableName">The name of the table to find.</param>
        /// <returns>The standalone table if found; otherwise, null.</returns>
        internal StandaloneTable FindTable(string tableName, Map targetMap = null)
        {
            // Check there is an input table name.
            if (string.IsNullOrEmpty(tableName))
                return null;

            // Use provided map or default to _activeMap.
            Map mapToUse = targetMap ?? _activeMap;

            try
            {
                // Finds tables by name and returns a read only list of standalone tables.
                IEnumerable<StandaloneTable> tables = mapToUse.FindStandaloneTables(tableName).OfType<StandaloneTable>();

                while (tables.Any())
                {
                    // Get the first table found by name.
                    StandaloneTable table = tables.First();

                    // Check the table is in the active map.
                    if (table.Map.Name.Equals(mapToUse.Name, StringComparison.OrdinalIgnoreCase))
                        return table;
                }
            }
            catch (Exception ex)
            {
                // Log the exception and return null.
                TraceLog($"FindTable error: Exception occurred while finding table. Table: {tableName}, Exception: {ex.Message}");
                return null;
            }

            return null;
        }

        /// <summary>
        /// Remove a table from the active map.
        /// </summary>
        /// <param name="tableName">The name of the table to remove.</param>
        /// <returns>The task result is true if successful; otherwise false.</returns>
        public async Task<bool> RemoveTableAsync(string tableName, Map targetMap = null)
        {
            // Check there is an input table name.
            if (string.IsNullOrEmpty(tableName))
                return false;

            // Use provided map or default to _activeMap.
            Map mapToUse = targetMap ?? _activeMap;

            try
            {
                // Find the table in the active map.
                StandaloneTable table = FindTable(tableName, mapToUse);

                if (table != null)
                {
                    // Remove the table.
                    await RemoveTableAsync(table, mapToUse);
                }

                return true;
            }
            catch (Exception ex)
            {
                // Log the exception and return false.
                TraceLog($"RemoveTableAsync error: Exception occurred while removing table. Table: {tableName}, Exception {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Remove a standalone table from the active map.
        /// </summary>
        /// <param name="table">The standalone table to remove.</param>
        /// <returns>The task result is true if successful; otherwise false.</returns>
        public async Task<bool> RemoveTableAsync(StandaloneTable table, Map targetMap = null)
        {
            // Check there is an input table name.
            if (table == null)
                return false;

            // Use provided map or default to _activeMap.
            Map mapToUse = targetMap ?? _activeMap;

            try
            {
                await QueuedTask.Run(() =>
                {
                    // Remove the table.
                    mapToUse.RemoveStandaloneTable(table);
                });
            }
            catch (Exception ex)
            {
                // Access to table.Name must occur on the CIM thread.
                string safeTableName = await QueuedTask.Run(() => table.Name);

                // Log the exception and return false.
                TraceLog($"RemoveTableAsync error: Exception occurred while removing table. Table: {safeTableName}, Exception: {ex.Message}");
                return false;
            }

            return true;
        }

        #endregion Tables

        #region Export

        /// <summary>
        /// Copy a feature class to a text fiile.
        /// </summary>
        /// <param name="inputLayer">The name of the input feature layer.</param>
        /// <param name="outFile">The path to the output text file.</param>
        /// <param name="columns">The list of columns to output (comma separated).</param>
        /// <param name="orderByColumns">The list of columns to order the output by (comma separated).</param>
        /// <param name="separator">The column separator.</param>
        /// <param name="append">Whether to append to the output file.</param>
        /// <param name="includeHeader">Whether to include a header row.</param>
        /// <param name="targetMap">Optional map to search in; defaults to the active map.</param>
        /// <returns>The task result is the number of rows written; -1 if error.</returns>
        public async Task<int> CopyFCToTextFileAsync(string inputLayer, string outFile, string columns, string orderByColumns,
             string separator, bool append = false, bool includeHeader = true, Map targetMap = null)
        {
            // Check there is an input layer name.
            if (string.IsNullOrEmpty(inputLayer))
                return -1;

            // Check there is an output table name.
            if (string.IsNullOrEmpty(outFile))
                return -1;

            // Check there are columns to output.
            if (string.IsNullOrEmpty(columns))
                return -1;

            bool missingColumns = false;
            string outColumns;
            FeatureLayer inputFeaturelayer;
            List<string> outColumnsList = [];
            List<string> orderByColumnsList = [];
            IReadOnlyList<Field> inputfields;

            try
            {
                // Get the input feature layer.
                inputFeaturelayer = await FindLayerAsync(inputLayer, targetMap);

                if (inputFeaturelayer == null)
                    return -1;

                // Get the list of fields for the input table.
                inputfields = await GetFCFieldsAsync(inputLayer, targetMap);

                // Check a list of fields is returned.
                if (inputfields == null || inputfields.Count == 0)
                    return -1;

                // Align the columns with what actually exists in the layer.
                List<string> columnsList = [.. columns.Split(',')];
                outColumns = "";
                foreach (string column in columnsList)
                {
                    string columnName = column.Trim();
                    if ((columnName.Substring(0, 1) == "\"") || (FieldExists(inputfields, columnName)))
                    {
                        outColumnsList.Add(columnName);
                        outColumns = outColumns + columnName + separator;
                    }
                    else
                    {
                        missingColumns = true;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                // Log the exception and return -1.
                TraceLog($"CopyFCToTextFileAsync error: Exception occurred while copying feature class to text file. Layer: {inputLayer}, OutFile: {outFile}, Exception: {ex.Message}");
                return -1;
            }

            // Stop if there aren't any columns.
            if (outColumnsList.Count == 0 || string.IsNullOrEmpty(outColumns))
                return -1;

            // Stop if there are any missing columns.
            if (missingColumns || string.IsNullOrEmpty(columns))
                return -1;

            // Remove the final separator.
            outColumns = outColumns[..^1];

            // Open output file.
            using StreamWriter txtFile = new(outFile, append);

            // Write the header if required.
            if (!append && includeHeader)
                txtFile.WriteLine(outColumns);

            int intLineCount = 0;
            try
            {
                await QueuedTask.Run(() =>
                {
                    /// Get the feature class for the input feature layer.
                    using FeatureClass featureClass = inputFeaturelayer.GetFeatureClass();

                    // Get the feature class defintion.
                    using FeatureClassDefinition featureClassDefinition = featureClass.GetDefinition();

                    // Create a row cursor.
                    RowCursor rowCursor;

                    // Create a new list of sort descriptions.
                    List<SortDescription> sortDescriptions = [];

                    if (!string.IsNullOrEmpty(orderByColumns))
                    {
                        orderByColumnsList = [.. orderByColumns.Split(',')];

                        // Build the list of sort descriptions for each orderby column in the input layer.
                        foreach (string column in orderByColumnsList)
                        {
                            // Get the column name (ignoring any trailing ASC/DESC sort order).
                            string columnName = column.Trim();
                            if (columnName.Contains(' '))
                                columnName = columnName.Split(" ")[0].Trim();

                            // Set the sort order to ascending or descending.
                            SortOrder sortOrder = SortOrder.Ascending;
                            if ((column.EndsWith(" DES", true, System.Globalization.CultureInfo.CurrentCulture)) ||
                               (column.EndsWith(" DESC", true, System.Globalization.CultureInfo.CurrentCulture)))
                                sortOrder = SortOrder.Descending;

                            // If the column is in the input table use it for sorting.
                            if ((columnName.Substring(0, 1) != "\"") && (FieldExists(inputfields, columnName)))
                            {
                                // Get the field from the feature class definition.
                                using Field field = featureClassDefinition.GetFields()
                                  .First(x => x.Name.Equals(columnName, StringComparison.OrdinalIgnoreCase));

                                // Create a SortDescription for the field.
                                SortDescription sortDescription = new(field)
                                {
                                    CaseSensitivity = CaseSensitivity.Insensitive,
                                    SortOrder = sortOrder
                                };

                                // Add the SortDescription to the list.
                                sortDescriptions.Add(sortDescription);
                            }
                        }

                        // Create a TableSortDescription.
                        TableSortDescription tableSortDescription = new(sortDescriptions);

                        // Create a cursor of the sorted features.
                        rowCursor = featureClass.Sort(tableSortDescription);
                    }
                    else
                    {
                        // Create a cursor of the features.
                        rowCursor = featureClass.Search();
                    }

                    // Loop through the feature class/table using the cursor.
                    while (rowCursor.MoveNext())
                    {
                        // Get the current row.
                        using Row record = rowCursor.Current;

                        string newRow = "";
                        foreach (string column in outColumnsList)
                        {
                            string columnName = column.Trim();

                            // If the column name isn't a literal.
                            if (columnName.Substring(0, 1) != "\"")
                            {
                                // Get the field value.
                                var columnValue = record[columnName];
                                columnValue ??= "";

                                // Wrap value if quotes if it is a string that contains a comma
                                if ((columnValue is string) && (columnValue.ToString().Contains(',')))
                                    columnValue = "\"" + columnValue.ToString() + "\"";

                                // Append the column value to the new row.
                                newRow = newRow + columnValue.ToString() + separator;
                            }
                            else
                            {
                                // Append the literal to the new row.
                                newRow = newRow + columnName + separator;
                            }
                        }

                        // Remove the final separator.
                        newRow = newRow[..^1];

                        // Write the new row.
                        txtFile.WriteLine(newRow);
                        intLineCount++;
                    }

                    // Dispose of the objects.
                    rowCursor.Dispose();
                    rowCursor = null;
                });
            }
            catch (Exception ex)
            {
                // Log the exception and return -1.
                TraceLog($"CopyFCToTextFileAsync error: Exception occurred while copying feature class to text file. Layer: {inputLayer}, OutFile: {outFile}, Exception: {ex.Message}");
                return -1;
            }
            finally
            {
                // Close the file.
                txtFile.Close();

                // Dispose of the object.
                txtFile.Dispose();
            }

            return intLineCount;
        }

        /// <summary>
        /// Copy a table to a text file.
        /// </summary>
        /// <param name="inputLayer">The name of the input table.</param>
        /// <param name="outFile">The path to the output text file.</param>
        /// <param name="columns">  The list of columns to output (comma separated).</param>
        /// <param name="orderByColumns">The list of columns to order the output by (comma separated).</param>
        /// <param name="separator">The column separator.</param>
        /// <param name="append">Whether to append to the output file.</param>
        /// <param name="includeHeader">Whether to include a header row.</param>
        /// <param name="targetMap">Optional map to search in; defaults to the active map.</param>
        /// <returns>The task result is the number of rows written; -1 if error.</returns>
        public async Task<int> CopyTableToTextFileAsync(string inputLayer, string outFile, string columns, string orderByColumns,
            string separator, bool append = false, bool includeHeader = true, Map targetMap = null)
        {
            // Check there is an input table name.
            if (string.IsNullOrEmpty(inputLayer))
                return -1;

            // Check there is an output table name.
            if (string.IsNullOrEmpty(outFile))
                return -1;

            // Check there are columns to output.
            if (string.IsNullOrEmpty(columns))
                return -1;

            bool missingColumns = false;
            string outColumns;
            StandaloneTable inputTable;
            List<string> outColumnsList = [];
            List<string> orderByColumnsList = [];
            IReadOnlyList<Field> inputfields;

            try
            {
                // Get the input feature layer.
                inputTable = FindTable(inputLayer, targetMap);

                if (inputTable == null)
                    return -1;

                // Get the list of fields for the input table.
                inputfields = await GetTableFieldsAsync(inputLayer, targetMap);

                // Check a list of fields is returned.
                if (inputfields == null || inputfields.Count == 0)
                    return -1;

                // Align the columns with what actually exists in the layer.
                List<string> columnsList = [.. columns.Split(',')];
                outColumns = "";
                foreach (string column in columnsList)
                {
                    string columnName = column.Trim();
                    if ((columnName.Substring(0, 1) == "\"") || (FieldExists(inputfields, columnName)))
                    {
                        outColumnsList.Add(columnName);
                        outColumns = outColumns + columnName + separator;
                    }
                    else
                    {
                        missingColumns = true;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                // Log the exception and return -1.
                TraceLog($"CopyTableToTextFileAsync error: Exception occurred while copying table to text file. Layer: {inputLayer}, OutFile: {outFile}, Exception: {ex.Message}");
                return -1;
            }

            // Stop if there aren't any columns.
            if (outColumnsList.Count == 0 || string.IsNullOrEmpty(outColumns))
                return -1;

            // Stop if there are any missing columns.
            if (missingColumns || string.IsNullOrEmpty(columns))
                return -1;

            // Remove the final separator.
            outColumns = outColumns[..^1];

            // Open output file.
            using StreamWriter txtFile = new(outFile, append);

            // Write the header if required.
            if (!append && includeHeader)
                txtFile.WriteLine(outColumns);

            int intLineCount = 0;
            try
            {
                await QueuedTask.Run(() =>
                {
                    /// Get the underlying table for the input layer.
                    using Table table = inputTable.GetTable();

                    // Get the table defintion.
                    using TableDefinition tableDefinition = table.GetDefinition();

                    // Create a row cursor.
                    RowCursor rowCursor;

                    // Create a new list of sort descriptions.
                    List<SortDescription> sortDescriptions = [];

                    if (!string.IsNullOrEmpty(orderByColumns))
                    {
                        orderByColumnsList = [.. orderByColumns.Split(',')];

                        // Build the list of sort descriptions for each orderby column in the input layer.
                        foreach (string column in orderByColumnsList)
                        {
                            // Get the column name (ignoring any trailing ASC/DESC sort order).
                            string columnName = column.Trim();
                            if (columnName.Contains(' '))
                                columnName = columnName.Split(" ")[0].Trim();

                            // Set the sort order to ascending or descending.
                            SortOrder sortOrder = SortOrder.Ascending;
                            if ((column.EndsWith(" DES", true, System.Globalization.CultureInfo.CurrentCulture)) ||
                               (column.EndsWith(" DESC", true, System.Globalization.CultureInfo.CurrentCulture)))
                                sortOrder = SortOrder.Descending;

                            // If the column is in the input table use it for sorting.
                            if ((columnName.Substring(0, 1) != "\"") && (FieldExists(inputfields, columnName)))
                            {
                                // Get the field from the feature class definition.
                                using Field field = tableDefinition.GetFields()
                                  .First(x => x.Name.Equals(columnName, StringComparison.OrdinalIgnoreCase));

                                // Create a SortDescription for the field.
                                SortDescription sortDescription = new(field)
                                {
                                    CaseSensitivity = CaseSensitivity.Insensitive,
                                    SortOrder = sortOrder
                                };

                                // Add the SortDescription to the list.
                                sortDescriptions.Add(sortDescription);
                            }
                        }

                        // Create a TableSortDescription.
                        TableSortDescription tableSortDescription = new(sortDescriptions);

                        // Create a cursor of the sorted features.
                        rowCursor = table.Sort(tableSortDescription);
                    }
                    else
                    {
                        // Create a cursor of the features.
                        rowCursor = table.Search();
                    }

                    // Loop through the feature class/table using the cursor.
                    while (rowCursor.MoveNext())
                    {
                        // Get the current row.
                        using Row record = rowCursor.Current;

                        string newRow = "";
                        foreach (string column in outColumnsList)
                        {
                            string columnName = column.Trim();

                            // If the column name isn't a literal.
                            if (columnName.Substring(0, 1) != "\"")
                            {
                                // Get the field value.
                                var columnValue = record[columnName];
                                columnValue ??= "";

                                // Wrap value if quotes if it is a string that contains a comma
                                if ((columnValue is string) && (columnValue.ToString().Contains(',')))
                                    columnValue = "\"" + columnValue.ToString() + "\"";

                                // Append the column value to the new row.
                                newRow = newRow + columnValue.ToString() + separator;
                            }
                            else
                            {
                                // Append the literal to the new row.
                                newRow = newRow + columnName + separator;
                            }
                        }

                        // Remove the final separator.
                        newRow = newRow[..^1];

                        // Write the new row.
                        txtFile.WriteLine(newRow);
                        intLineCount++;
                    }

                    // Dispose of the objects.
                    rowCursor.Dispose();
                    rowCursor = null;
                });
            }
            catch (Exception ex)
            {
                // Log the exception and return -1.
                TraceLog($"CopyTableToTextFileAsync error: Exception occurred while copying table to text file. Layer: {inputLayer}, OutFile: {outFile}, Exception {ex.Message}");
                return -1;
            }
            finally
            {
                // Close the file.
                txtFile.Close();

                // Dispose of the object.
                txtFile.Dispose();
            }

            return intLineCount;
        }

        /// <summary>
        /// Copy a table to a text file.
        /// </summary>
        /// <param name="inTable">The name of the input table.</param>
        /// <param name="outFile">The path to the output text file.</param>
        /// <param name="isSpatial">Whether the input table is spatial (feature class) or non-spatial (table).</param>
        /// <param name="append">Whether to append to the output file.</param>
        /// <returns>The task result is the number of rows written; -1 if error.</returns>
        public async Task<int> CopyToCSVAsync(string inTable, string outFile, bool isSpatial, bool append)
        {
            // Check if there is an input table name.
            if (string.IsNullOrEmpty(inTable))
                return -1;

            // Check if there is an output file.
            if (string.IsNullOrEmpty(outFile))
                return -1;

            string separator = ",";
            return await CopyToTextFileAsync(inTable, outFile, separator, isSpatial, append);
        }

        /// <summary>
        /// Copy a table to a text file.
        /// </summary>
        /// <param name="inTable">The name of the input table.</param>
        /// <param name="outFile">The path to the output text file.</param>
        /// <param name="isSpatial">Whether the input table is spatial (feature class) or non-spatial (table).</param>
        /// <param name="append">Whether to append to the output file.</param>
        /// <returns>The task result is the number of rows written; -1 if error.</returns>
        public async Task<int> CopyToTabAsync(string inTable, string outFile, bool isSpatial, bool append)
        {
            // Check if there is an input table name.
            if (string.IsNullOrEmpty(inTable))
                return -1;

            // Check if there is an output file.
            if (string.IsNullOrEmpty(outFile))
                return -1;

            string separator = "\t";
            return await CopyToTextFileAsync(inTable, outFile, separator, isSpatial, append);
        }

        /// <summary>
        /// Copy a table to a text file.
        /// </summary>
        /// <param name="inputLayer">The name of the input table.</param>
        /// <param name="outFile">The path to the output text file.</param>
        /// <param name="separator">The column separator.</param>
        /// <param name="isSpatial">Whether the input table is spatial (feature class) or non-spatial (table).</param>
        /// <param name="append">Whether to append to the output file.</param>
        /// <param name="includeHeader">Whether to include a header row.</param>
        /// <returns>The task result is the number of rows written; -1 if error.</returns>
        public async Task<int> CopyToTextFileAsync(string inputLayer, string outFile, string separator, bool isSpatial, bool append = false,
            bool includeHeader = true, Map targetMap = null)
        {
            // Check there is an input table name.
            if (string.IsNullOrEmpty(inputLayer))
                return -1;

            // Check there is an output file.
            if (string.IsNullOrEmpty(outFile))
                return -1;

            string fieldName = null;
            string header = "";
            int ignoreField = -1;

            int intFieldCount;
            try
            {
                IReadOnlyList<Field> fields;

                if (isSpatial)
                {
                    // Get the list of fields for the input table.
                    fields = await GetFCFieldsAsync(inputLayer, targetMap);
                }
                else
                {
                    // Get the list of fields for the input table.
                    fields = await GetTableFieldsAsync(inputLayer, targetMap);
                }

                // Check a list of fields is returned.
                if (fields == null || fields.Count == 0)
                    return -1;

                intFieldCount = fields.Count;

                // Iterate through the fields in the collection to create header
                // and flag which fields to ignore.
                for (int i = 0; i < intFieldCount; i++)
                {
                    // Get the fieldName name.
                    fieldName = fields[i].Name;

                    using Field field = fields[i];

                    // Get the fieldName type.
                    FieldType fieldType = field.FieldType;

                    string fieldTypeName = fieldType.ToString();

                    if (fieldName.Equals("sp_geometry", StringComparison.OrdinalIgnoreCase) || fieldName.Equals("shape", StringComparison.OrdinalIgnoreCase))
                        ignoreField = i;
                    else
                        header = header + fieldName + separator;
                }

                if (!append && includeHeader)
                {
                    // Remove the final separator from the header.
                    header = header.Substring(0, header.Length - 1);

                    // Write the header to the output file.
                    FileFunctions.WriteEmptyTextFile(outFile, header);
                }
            }
            catch (Exception ex)
            {
                // Log the exception and return -1.
                TraceLog($"CopyToTextFileAsync error: Exception occurred while copying table to text file. Layer: {inputLayer}, OutFile: {outFile}, Exception: {ex.Message}");
                return -1;
            }

            // Open output file.
            StreamWriter txtFile = new(outFile, append);

            int intLineCount = 0;
            try
            {
                await QueuedTask.Run(async () =>
                {
                    // Create a row cursor.
                    RowCursor rowCursor;

                    if (isSpatial)
                    {
                        FeatureLayer inputFC;

                        // Get the input feature layer.
                        inputFC = await FindLayerAsync(inputLayer, targetMap);

                        /// Get the underlying table for the input layer.
                        using FeatureClass featureClass = inputFC.GetFeatureClass();

                        // Create a cursor of the features.
                        rowCursor = featureClass.Search();
                    }
                    else
                    {
                        StandaloneTable inputTable;

                        // Get the input table.
                        inputTable = FindTable(inputLayer, targetMap);

                        /// Get the underlying table for the input layer.
                        using Table table = inputTable.GetTable();

                        // Create a cursor of the features.
                        rowCursor = table.Search();
                    }

                    // Loop through the feature class/table using the cursor.
                    while (rowCursor.MoveNext())
                    {
                        // Get the current row.
                        using Row row = rowCursor.Current;

                        // Loop through the fields.
                        string rowStr = "";
                        for (int i = 0; i < intFieldCount; i++)
                        {
                            // String the column values together (if they are not to be ignored).
                            if (i != ignoreField)
                            {
                                // Get the column value.
                                var colValue = row.GetOriginalValue(i);

                                // Wrap the value if quotes if it is a string that contains a comma
                                string colStr = null;
                                if (colValue != null)
                                {
                                    if ((colValue is string) && (colValue.ToString().Contains(',')))
                                        colStr = "\"" + colValue.ToString() + "\"";
                                    else
                                        colStr = colValue.ToString();
                                }

                                // Add the column string to the row string.
                                rowStr += colStr;

                                // Add the column separator (if not the last column).
                                if (i < intFieldCount - 1)
                                    rowStr += separator;
                            }
                        }

                        // Write the row string to the output file.
                        txtFile.WriteLine(rowStr);
                        intLineCount++;
                    }
                    // Dispose of the objects.
                    rowCursor.Dispose();
                    rowCursor = null;
                });
            }
            catch (Exception ex)
            {
                // Log the exception and return -1.
                TraceLog($"CopyToTextFileAsync error: Exception occurred while copying table to text file. Layer: {inputLayer}, OutFile: {outFile}, Exception: {ex.Message}");
                return -1;
            }
            finally
            {
                // Close the output file and dispose of the object.
                txtFile.Close();
                txtFile.Dispose();
            }

            return intLineCount;
        }

        #endregion Export
    }
}
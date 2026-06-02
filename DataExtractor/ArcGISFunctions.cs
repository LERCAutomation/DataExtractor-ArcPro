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

using ArcGIS.Core.Data;
using ArcGIS.Core.Data.DDL;
using ArcGIS.Core.Data.Exceptions;
using ArcGIS.Desktop.Catalog;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Geoprocessing;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using QueryFilter = ArcGIS.Core.Data.QueryFilter;

namespace DataTools
{
    /// <summary>
    /// This helper class provides ArcGIS Pro feature class and layer functions.
    /// </summary>
    internal static class ArcGISFunctions
    {
        #region Feature Class

        /// <summary>
        /// Check if the feature class exists in the file path.
        /// </summary>
        /// <param name="filePath">The file path to the geodatabase.</param>
        /// <param name="fileName">The feature class name.</param>
        /// <returns>True if the feature class exists, false if not.</returns>
        public static async Task<bool> FeatureClassExistsAsync(string filePath, string fileName)
        {
            // Check there is an input file path.
            if (string.IsNullOrEmpty(filePath))
                return false;

            // Check there is an input file name.
            if (string.IsNullOrEmpty(fileName))
                return false;

            if (fileName.Substring(fileName.Length - 4, 1) == ".")
            {
                // It's a file.
                if (FileFunctions.FileExists(filePath + @"\" + fileName))
                    return true;
                else
                    return false;
            }
            else if (filePath.Substring(filePath.Length - 3, 3).Equals("sde", StringComparison.OrdinalIgnoreCase))
            {
                // It's an SDE class. Not handled (use SQL Server Functions).
                return false;
            }
            else // It is a geodatabase class.
            {
                try
                {
                    return await FeatureClassExistsGDBAsync(filePath, fileName);
                }
                catch
                {
                    // GetDefinition throws an exception if the definition doesn't exist.
                    return false;
                }
            }
        }

        /// <summary>
        /// Check if the feature class exists.
        /// </summary>
        /// <param name="fullPath">The full file path to the feature class.</param>
        /// <returns>True if the feature class exists, false if not.</returns>
        public static async Task<bool> FeatureClassExistsAsync(string fullPath)
        {
            // Check there is an input file path.
            if (string.IsNullOrEmpty(fullPath))
                return false;

            return await FeatureClassExistsAsync(FileFunctions.GetDirectoryName(fullPath), FileFunctions.GetFileName(fullPath));
        }

        /// <summary>
        /// Delete a feature class by file path and file name.
        /// </summary>
        /// <param name="filePath">The file path to the geodatabase.</param>
        /// <param name="fileName">The feature class name.</param>
        /// <returns>True if the feature class was deleted, false if not.</returns>
        public static async Task<bool> DeleteFeatureClassAsync(string filePath, string fileName)
        {
            // Check there is an input file path.
            if (string.IsNullOrEmpty(filePath))
                return false;

            // Check there is an input file name.
            if (string.IsNullOrEmpty(fileName))
                return false;

            string featureClass = filePath + @"\" + fileName;

            return await DeleteFeatureClassAsync(featureClass);
        }

        /// <summary>
        /// Delete a feature class by file name.
        /// </summary>
        /// <param name="fileName">The full file path to the feature class.</param>
        /// <returns>True if the feature class was deleted, false if not.</returns>
        public static async Task<bool> DeleteFeatureClassAsync(string fileName)
        {
            // Check there is an input file name.
            if (string.IsNullOrEmpty(fileName))
                return false;

            // Make a value array of strings to be passed to the tool.
            var parameters = Geoprocessing.MakeValueArray(fileName);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geoprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; //| GPExecuteToolFlags.RefreshProjectItems;

            //Geoprocessing.OpenToolDialog("management.Delete", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.Delete", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        /// <summary>
        /// Add a field to a feature class or table.
        /// </summary>
        /// <param name="inTable">The input table or feature class.</param>
        /// <param name="fieldName">The name of the field to add.</param>
        /// <param name="fieldType">The type of field to add.</param>
        /// <param name="fieldPrecision">The precision of the field to add.</param>
        /// <param name="fieldScale">The scale of the field to add.</param>
        /// <param name="fieldLength">The length of the field to add.</param>
        /// <param name="fieldAlias">The alias of the field to add.</param>
        /// <param name="fieldIsNullable">Whether the field is nullable.</param>
        /// <param name="fieldIsRequred">Whether the field is required.</param>
        /// <param name="fieldDomain">The domain of the field to add.</param>
        /// <returns>True if the field was added, false if not.</returns>
        public static async Task<bool> AddFieldAsync(string inTable, string fieldName, string fieldType = "TEXT",
            long fieldPrecision = -1, long fieldScale = -1, long fieldLength = -1, string fieldAlias = null,
            bool fieldIsNullable = true, bool fieldIsRequred = false, string fieldDomain = null)
        {
            // Check if there is an input table name.
            if (string.IsNullOrEmpty(inTable))
                return false;

            // Check if there is an input field name.
            if (string.IsNullOrEmpty(fieldName))
                return false;

            // Make a value array of strings to be passed to the tool.
            var parameters = Geoprocessing.MakeValueArray(inTable, fieldName, fieldType,
                fieldPrecision > 0 ? fieldPrecision : null, fieldScale > 0 ? fieldScale : null, fieldLength > 0 ? fieldLength : null,
                fieldAlias ?? null, fieldIsNullable ? "NULLABLE" : "NON_NULLABLE",
                fieldIsRequred ? "REQUIRED" : "NON_REQUIRED", fieldDomain);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geoprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; //| GPExecuteToolFlags.RefreshProjectItems;

            //Geoprocessing.OpenToolDialog("management.AddField", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.AddField", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        /// <summary>
        /// Rename a field in a feature class or table.
        /// </summary>
        /// <param name="inTable">The input table or feature class.</param>
        /// <param name="fieldName">The name of the field to rename.</param>
        /// <param name="newFieldName">The new name of the field.</param>
        /// <returns>True if the field was renamed, false if not.</returns>
        public static async Task<bool> RenameFieldAsync(string inTable, string fieldName, string newFieldName)
        {
            // Check if there is an input table name.
            if (string.IsNullOrEmpty(inTable))
                return false;

            // Check if there is an input old field name.
            if (string.IsNullOrEmpty(fieldName))
                return false;

            // Check if there is an input new field name.
            if (string.IsNullOrEmpty(newFieldName))
                return false;

            // Make a value array of strings to be passed to the tool.
            var parameters = Geoprocessing.MakeValueArray(inTable, fieldName, newFieldName);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geoprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; //| GPExecuteToolFlags.RefreshProjectItems;

            //Geoprocessing.OpenToolDialog("management.AlterField", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.AlterField", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        /// <summary>
        /// Calculate a field in a feature class or table.
        /// </summary>
        /// <param name="inTable">The input table or feature class.</param>
        /// <param name="fieldName">The name of the field to calculate.</param>
        /// <param name="fieldCalc">The calculation string for the field.</param>
        /// <returns>True if the field was calculated, false if not.</returns>
        public static async Task<bool> CalculateFieldAsync(string inTable, string fieldName, string fieldCalc)
        {
            // Check if there is an input table name.
            if (string.IsNullOrEmpty(inTable))
                return false;

            // Check if there is an input field name.
            if (string.IsNullOrEmpty(fieldName))
                return false;

            // Check if there is an input field calculcation string.
            if (string.IsNullOrEmpty(fieldCalc))
                return false;

            // Make a value array of strings to be passed to the tool.
            var parameters = Geoprocessing.MakeValueArray(inTable, fieldName, fieldCalc);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geoprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; //| GPExecuteToolFlags.RefreshProjectItems;

            //Geoprocessing.OpenToolDialog("management.CalculateField", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.CalculateField", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        /// <summary>
        /// Calculate the geometry of a feature class.
        /// </summary>
        /// <param name="inTable">The input table or feature class.</param>
        /// <param name="geometryProperty">The geometry property to calculate.</param>
        /// <param name="lineUnit">The line unit to use.</param>
        /// <param name="areaUnit">The area unit to use.</param>
        /// <returns>True if the geometry was calculated, false if not.</returns>
        public static async Task<bool> CalculateGeometryAsync(string inTable, string geometryProperty, string lineUnit = "", string areaUnit = "")
        {
            // Check if there is an input table name.
            if (string.IsNullOrEmpty(inTable))
                return false;

            // Check if there is an input geometry property.
            if (string.IsNullOrEmpty(geometryProperty))
                return false;

            // Make a value array of strings to be passed to the tool.
            var parameters = Geoprocessing.MakeValueArray(inTable, geometryProperty, lineUnit, areaUnit);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geoprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; //| GPExecuteToolFlags.RefreshProjectItems;

            //Geoprocessing.OpenToolDialog("management.CalculateGeometryAttributes", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.CalculateGeometryAttributes", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        /// <summary>
        /// Count the features in a layer using a search where clause.
        /// </summary>
        /// <param name="layer">The input feature layer.</param>
        /// <param name="whereClause">The where clause to use.</param>
        /// <param name="subfields">The subfields to use.</param>
        /// <param name="prefixClause">The prefix clause to use.</param>
        /// <param name="postfixClause">The postfix clause to use.</param>
        /// <returns>long</returns>
        public static async Task<long> GetFeaturesCountAsync(FeatureLayer layer, string whereClause = null, string subfields = null, string prefixClause = null, string postfixClause = null)
        {
            // Check if there is an input layer name.
            if (layer == null)
                return -1;

            long featureCount = 0;
            try
            {
                // Create a query filter using the where clause.
                QueryFilter queryFilter = new();

                // Apply where clause.
                if (!string.IsNullOrEmpty(whereClause))
                    queryFilter.WhereClause = whereClause;

                // Apply subfields clause.
                if (!string.IsNullOrEmpty(subfields))
                    queryFilter.SubFields = subfields;

                // Apply prefix clause.
                if (!string.IsNullOrEmpty(prefixClause))
                    queryFilter.PrefixClause = prefixClause;

                // Apply postfix clause.
                if (!string.IsNullOrEmpty(postfixClause))
                    queryFilter.PostfixClause = postfixClause;

                await QueuedTask.Run(() =>
                {
                    /// Count the number of features matching the search clause.
                    using FeatureClass featureClass = layer.GetFeatureClass();

                    featureCount = featureClass.GetCount(queryFilter);
                });
            }
            catch
            {
                // Handle Exception.
                return -1;
            }

            return featureCount;
        }

        /// <summary>
        /// Count the duplicate features in a layer using a search where clause.
        /// </summary>
        /// <param name="layer">The input feature layer.</param>
        /// <param name="keyField">The key field to check for duplicates.</param>
        /// <param name="whereClause">The where clause to use.</param>
        /// <returns>long</returns>
        public static async Task<long> GetDuplicateFeaturesCountAsync(FeatureLayer layer, string keyField, string whereClause = null)
        {
            // Check if there is an input layer name.
            if (layer == null)
                return -1;

            // Check if there is a input key field.
            if (string.IsNullOrEmpty(keyField))
                return -1;

            long featureCount = 0;
            try
            {
                // Create a query filter using the where clause.
                QueryFilter queryFilter = new();

                // Apply where clause.
                if (!string.IsNullOrEmpty(whereClause))
                    queryFilter.WhereClause = whereClause;

                // Apply subfields clause.
                if (!string.IsNullOrEmpty(keyField))
                    queryFilter.SubFields = keyField;

                List<string> keys = [];

                await QueuedTask.Run(() =>
                {
                    /// Get the feature class for the layer.
                    using FeatureClass featureClass = layer.GetFeatureClass();

                    // Create a cursor of the features.
                    using RowCursor rowCursor = featureClass.Search(queryFilter);

                    // Loop through the feature class/table using the cursor.
                    while (rowCursor.MoveNext())
                    {
                        // Get the current row.
                        using Row record = rowCursor.Current;

                        // Get the key value.
                        string key = Convert.ToString(record[keyField]);
                        key ??= "";

                        // Add the key to the list of keys.
                        keys.Add(key);
                    }
                    // Dispose of the objects.
                    featureClass.Dispose();
                    rowCursor.Dispose();

                    // Get a list of any duplicate keys.
                    List<string> duplicateKeys = [.. keys.GroupBy(x => x)
                      .Where(g => g.Count() > 1)
                      .Select(y => y.Key)];

                    // Return how many duplicate keys there are.
                    featureCount = duplicateKeys.Count;
                });
            }
            catch
            {
                // Handle Exception.
                return -1;
            }

            return featureCount;
        }

        /// <summary>
        /// Buffer the features in a feature class with a specified distance.
        /// </summary>
        /// <param name="inFeatureClass">The input feature class.</param>
        /// <param name="outFeatureClass">The output feature class.</param>
        /// <param name="bufferDistance">The buffer distance.</param>
        /// <param name="lineSide">The line side option.</param>
        /// <param name="lineEndType">The line end type option.</param>
        /// <param name="dissolveOption">The dissolve option.</param>
        /// <param name="dissolveFields">The dissolve fields.</param>
        /// <param name="method">The method option.</param>
        /// <param name="addToMap">Whether to add the output to the map.</param>
        /// <returns>True if the features were buffered, false if not.</returns>
        public static async Task<bool> BufferFeaturesAsync(string inFeatureClass, string outFeatureClass, string bufferDistance,
            string lineSide = "FULL", string lineEndType = "ROUND", string dissolveOption = "NONE", string dissolveFields = "", string method = "PLANAR", bool addToMap = false)
        {
            // Check if there is an input feature class.
            if (string.IsNullOrEmpty(inFeatureClass))
                return false;

            // Check if there is an output feature class.
            if (string.IsNullOrEmpty(outFeatureClass))
                return false;

            // Check if there is an input buffer distance.
            if (string.IsNullOrEmpty(bufferDistance))
                return false;

            // Make a value array of strings to be passed to the tool.
            List<string> parameters = [.. Geoprocessing.MakeValueArray(inFeatureClass, outFeatureClass, bufferDistance, lineSide, lineEndType, dissolveOption)];
            if (!string.IsNullOrEmpty(dissolveFields))
                parameters.Add(dissolveFields);
            parameters.Add(method);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geoprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; // | GPExecuteToolFlags.RefreshProjectItems;
            if (addToMap)
                executeFlags |= GPExecuteToolFlags.AddOutputsToMap;

            //Geoprocessing.OpenToolDialog("analysis.Buffer", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("analysis.Buffer", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        /// <summary>
        /// Clip the features in a feature class using a clip feature layer.
        /// </summary>
        /// <param name="inFeatureClass">The input feature class.</param>
        /// <param name="clipFeatureClass">The clip feature class.</param>
        /// <param name="outFeatureClass">The output feature class.</param>
        /// <param name="addToMap">Whether to add the output to the map.</param>
        /// <returns>True if the features were clipped, false if not.</returns>
        public static async Task<bool> ClipFeaturesAsync(string inFeatureClass, string clipFeatureClass, string outFeatureClass, bool addToMap = false)
        {
            // Check if there is an input feature class.
            if (string.IsNullOrEmpty(inFeatureClass))
                return false;

            // Check if there is an input clip feature class.
            if (string.IsNullOrEmpty(clipFeatureClass))
                return false;

            // Check if there is an output feature class.
            if (string.IsNullOrEmpty(outFeatureClass))
                return false;

            // Make a value array of strings to be passed to the tool.
            List<string> parameters = [.. Geoprocessing.MakeValueArray(inFeatureClass, clipFeatureClass, outFeatureClass)];

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geoprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; // | GPExecuteToolFlags.RefreshProjectItems;
            if (addToMap)
                executeFlags |= GPExecuteToolFlags.AddOutputsToMap;

            //Geoprocessing.OpenToolDialog("analysis.Clip", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("analysis.Clip", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        /// <summary>
        /// Intersect the features in a feature class with another feature class.
        /// </summary>
        /// <param name="inFeatures">The input feature class.</param>
        /// <param name="outFeatureClass">The output feature class.</param>
        /// <param name="joinAttributes">Which attributes to join.</param>
        /// <param name="outputType">The output type.</param>
        /// <param name="addToMap">Whether to add the output to the map.</param>
        /// <returns>True if the features were intersected, false if not.</returns>
        public static async Task<bool> IntersectFeaturesAsync(string inFeatures, string outFeatureClass, string joinAttributes = "ALL", string outputType = "INPUT", bool addToMap = false)
        {
            // Check if there is an input feature class.
            if (string.IsNullOrEmpty(inFeatures))
                return false;

            // Check if there is an output feature class.
            if (string.IsNullOrEmpty(outFeatureClass))
                return false;

            // Make a value array of strings to be passed to the tool.
            List<string> parameters = [.. Geoprocessing.MakeValueArray(inFeatures, outFeatureClass, joinAttributes, outputType)];

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geoprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; // | GPExecuteToolFlags.RefreshProjectItems;
            if (addToMap)
                executeFlags |= GPExecuteToolFlags.AddOutputsToMap;

            //Geoprocessing.OpenToolDialog("analysis.Intersect", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("analysis.Intersect", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        /// <summary>
        /// Spatially join a feature class with another feature class.
        /// </summary>
        /// <param name="targetFeatures">The target feature class.</param>
        /// <param name="joinFeatures">The join feature class.</param>
        /// <param name="outFeatureClass">The output feature class.</param>
        /// <param name="joinOperation">The join operation type; one to one or one to many.</param>
        /// <param name="joinType">The join type; keep all or keep matching.</param>
        /// <param name="fieldMapping">The field mapping string.</param>
        /// <param name="matchOption">The match option; intersect, within a distance, etc.</param>
        /// <param name="searchRadius">The search radius.</param>
        /// <param name="distanceField">The distance field name.</param>
        /// <param name="matchFields">The match fields string.</param>
        /// <param name="addToMap">Whether to add the output to the map.</param>
        /// <returns>True if the features were spatially joined, false if not.</returns>
        public static async Task<bool> SpatialJoinAsync(string targetFeatures, string joinFeatures, string outFeatureClass, string joinOperation = "JOIN_ONE_TO_ONE",
            string joinType = "KEEP_ALL", string fieldMapping = "", string matchOption = "INTERSECT", string searchRadius = "0", string distanceField = "",
            string matchFields = "", bool addToMap = false)
        {
            // Check if there is an input target feature class.
            if (string.IsNullOrEmpty(targetFeatures))
                return false;

            // Check if there is an input join feature class.
            if (string.IsNullOrEmpty(joinFeatures))
                return false;

            // Check if there is an output feature class.
            if (string.IsNullOrEmpty(outFeatureClass))
                return false;

            // Make a value array of strings to be passed to the tool.
            List<string> parameters = [.. Geoprocessing.MakeValueArray(targetFeatures, joinFeatures, outFeatureClass, joinOperation, joinType, fieldMapping,
                matchOption, searchRadius, distanceField, matchFields)];

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geoprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; // | GPExecuteToolFlags.RefreshProjectItems;
            if (addToMap)
                executeFlags |= GPExecuteToolFlags.AddOutputsToMap;

            //Geoprocessing.OpenToolDialog("analysis.SpatialJoin", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("analysis.SpatialJoin", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        /// <summary>
        /// Permanently join fields from one feature class to another feature class.
        /// </summary>
        /// <param name="inFeatures">The input feature class.</param>
        /// <param name="inField">The input field name.</param>
        /// <param name="joinFeatures">The join feature class.</param>
        /// <param name="joinField">The join field name.</param>
        /// <param name="fields">The fields to join.</param>
        /// <param name="fmOption">The field mapping option; NOT_USE_FM or USE_FM.</param>
        /// <param name="fieldMapping">The field mapping string.</param>
        /// <param name="indexJoinFields">The index join fields option; NO_INDEXES or INDEXES.</param>
        /// <param name="addToMap">Whether to add the output to the map.</param>
        /// <returns>True if the fields were joined, false if not.</returns>
        public static async Task<bool> JoinFieldsAsync(string inFeatures, string inField, string joinFeatures, string joinField,
            string fields = "", string fmOption = "NOT_USE_FM", string fieldMapping = "", string indexJoinFields = "NO_INDEXES",
            bool addToMap = false)
        {
            // Check if there is an input target feature class.
            if (string.IsNullOrEmpty(inFeatures))
                return false;

            // Check if there is an input field name.
            if (string.IsNullOrEmpty(inField))
                return false;

            // Check if there is a join feature class.
            if (string.IsNullOrEmpty(joinFeatures))
                return false;

            // Check if there is a join field name.
            if (string.IsNullOrEmpty(joinField))
                return false;

            // Make a value array of strings to be passed to the tool.
            List<string> parameters = [.. Geoprocessing.MakeValueArray(inFeatures, inField, joinFeatures, joinField, fields,
                fmOption, fieldMapping, indexJoinFields)];

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geoprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; // | GPExecuteToolFlags.RefreshProjectItems;
            if (addToMap)
                executeFlags |= GPExecuteToolFlags.AddOutputsToMap;

            //Geoprocessing.OpenToolDialog("management.JoinField", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.JoinField", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        /// <summary>
        /// Calculate the summary statistics for a feature class or table.
        /// </summary>
        /// <param name="inTable">The input table or feature class.</param>
        /// <param name="outTable">The output table.</param>
        /// <param name="statisticsFields">The statistics fields string.</param>
        /// <param name="caseFields">The case fields string.</param>
        /// <param name="concatenationSeparator">The concatenation separator string.</param>
        /// <param name="addToMap">Whether to add the output to the map.</param>
        /// <returns>True if the summary statistics were calculated, false if not.</returns>
        public static async Task<bool> CalculateSummaryStatisticsAsync(string inTable, string outTable, string statisticsFields,
            string caseFields = "", string concatenationSeparator = "", bool addToMap = false)
        {
            // Check if there is an input table name.
            if (string.IsNullOrEmpty(inTable))
                return false;

            // Check if there is an output table name.
            if (string.IsNullOrEmpty(outTable))
                return false;

            // Check if there is an input statistics fields string.
            if (string.IsNullOrEmpty(statisticsFields))
                return false;

            // Make a value array of strings to be passed to the tool.
            List<string> parameters = [.. Geoprocessing.MakeValueArray(inTable, outTable, statisticsFields, caseFields, concatenationSeparator)];

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geoprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; // | GPExecuteToolFlags.RefreshProjectItems;
            if (addToMap)
                executeFlags |= GPExecuteToolFlags.AddOutputsToMap;

            //Geoprocessing.OpenToolDialog("analysis.Statistics", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("analysis.Statistics", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        /// <summary>
        /// Convert the features in a feature class to a point feature class.
        /// </summary>
        /// <param name="inFeatureClass">The input feature class.</param>
        /// <param name="outFeatureClass">The output feature class.</param>
        /// <param name="pointLocation">The point location option; CENTROID or INSIDE.</param>
        /// <param name="addToMap">Whether to add the output to the map.</param>
        /// <returns>True if the features were converted to points, false if not.</returns>
        public static async Task<bool> FeatureToPointAsync(string inFeatureClass, string outFeatureClass, string pointLocation = "CENTROID", bool addToMap = false)
        {
            // Check if there is an input feature class.
            if (string.IsNullOrEmpty(inFeatureClass))
                return false;

            // Check if there is an output feature class.
            if (string.IsNullOrEmpty(outFeatureClass))
                return false;

            // Make a value array of strings to be passed to the tool.
            List<string> parameters = [.. Geoprocessing.MakeValueArray(inFeatureClass, outFeatureClass, pointLocation)];

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geoprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; // | GPExecuteToolFlags.RefreshProjectItems;
            if (addToMap)
                executeFlags |= GPExecuteToolFlags.AddOutputsToMap;

            //Geoprocessing.OpenToolDialog("management.FeatureToPoint", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.FeatureToPoint", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        /// <summary>
        /// Convert the features in a feature class to a point feature class.
        /// </summary>
        /// <param name="inFeatureClass">The input feature class.</param>
        /// <param name="nearFeatureClass">The near feature class.</param>
        /// <param name="searchRadius">The search radius.</param>
        /// <param name="location">The location option; NO_LOCATION or LOCATION.</param>
        /// <param name="angle">The angle option; NO_ANGLE or ANGLE.</param>
        /// <param name="method">The method option; PLANAR or GEODESIC.</param>
        /// <param name="fieldNames">The field names string.</param>
        /// <param name="distanceUnit">The distance unit.</param>
        /// <returns>True if the near analysis was successful, false if not.</returns>
        public static async Task<bool> NearAnalysisAsync(string inFeatureClass, string nearFeatureClass, string searchRadius = "",
            string location = "NO_LOCATION", string angle = "NO_ANGLE", string method = "PLANAR", string fieldNames = "", string distanceUnit = "")
        {
            // Check if there is an input feature class.
            if (string.IsNullOrEmpty(inFeatureClass))
                return false;

            // Check if there is an output feature class.
            if (string.IsNullOrEmpty(nearFeatureClass))
                return false;

            // Make a value array of strings to be passed to the tool.
            List<string> parameters = [.. Geoprocessing.MakeValueArray(inFeatureClass, nearFeatureClass, searchRadius, location, angle, method, fieldNames, distanceUnit)];

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geoprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; // | GPExecuteToolFlags.RefreshProjectItems;

            //Geoprocessing.OpenToolDialog("analysis.Near", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("analysis.Near", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        #endregion Feature Class

        #region Geodatabase

        /// <summary>
        /// Create a new file geodatabase.
        /// </summary>
        /// <param name="fullPath">The full path to the new file geodatabase.</param>
        /// <returns>Geodatabase</returns>
        public static Geodatabase CreateFileGeodatabase(string fullPath)
        {
            // Check if there is an input full path.
            if (string.IsNullOrEmpty(fullPath))
                return null;

            Geodatabase geodatabase;

            try
            {
                // Create a FileGeodatabaseConnectionPath with the name of the file geodatabase you wish to create
                FileGeodatabaseConnectionPath fileGeodatabaseConnectionPath = new(new Uri(fullPath));

                // Create and use the file geodatabase
                geodatabase = SchemaBuilder.CreateGeodatabase(fileGeodatabaseConnectionPath);
            }
            catch
            {
                // Handle Exception.
                return null;
            }

            return geodatabase;
        }

        /// <summary>
        /// Deletes a file geodatabase at the specified path, retrying if it's temporarily locked.
        /// </summary>
        /// <param name="fullPath">The full path to the .gdb folder to delete.</param>
        /// <returns>True if the geodatabase was successfully deleted; false otherwise.</returns>
        public static async Task<bool> DeleteFileGeodatabaseAsync(string fullPath)
        {
            // Check if there is an input full path.
            if (string.IsNullOrEmpty(fullPath))
                return false;

            bool success = false;

            // Try up to 5 times in case the geodatabase is temporarily locked.
            for (int attempt = 0; attempt < 5; attempt++)
            {
                try
                {
                    // Run the delete operation on the QueuedTask to ensure it's on the correct ArcGIS Pro thread.
                    await QueuedTask.Run(() =>
                    {
                        // Create a FileGeodatabaseConnectionPath using the full path
                        FileGeodatabaseConnectionPath fileGeodatabaseConnectionPath = new(new Uri(fullPath));

                        // Delete the file geodatabase using SchemaBuilder
                        SchemaBuilder.DeleteGeodatabase(fileGeodatabaseConnectionPath);
                    });

                    // If no exception was thrown, deletion was successful
                    success = true;
                    break;
                }
                catch (IOException)
                {
                    // Likely a file lock — wait briefly before retrying
                    await Task.Delay(2000);
                }
                catch (GeodatabaseNotFoundOrOpenedException)
                {
                    // GDB does not exist or is still open in ArcGIS Pro — not retryable
                    break;
                }
                catch (GeodatabaseTableException)
                {
                    // One or more tables may still be locked or open — not retryable
                    break;
                }
                catch (Exception)
                {
                    // Unexpected error — break to avoid silent failure
                    break;
                }
            }

            // Return whether the operation succeeded
            return success;
        }

        /// <summary>
        /// Check if a feature class exists in a geodatabase.
        /// </summary>
        /// <param name="filePath">The file path to the geodatabase.</param>
        /// <param name="fileName">The feature class name.</param>
        /// <returns>True if the feature class exists; false otherwise.</returns>
        public static async Task<bool> FeatureClassExistsGDBAsync(string filePath, string fileName)
        {
            // Check there is an input file path.
            if (string.IsNullOrEmpty(filePath))
                return false;

            // Check there is an input file name.
            if (string.IsNullOrEmpty(fileName))
                return false;

            bool exists = false;

            try
            {
                await QueuedTask.Run(() =>
                {
                    // Open the file geodatabase. This will open the geodatabase if the folder exists and contains a valid geodatabase.
                    using Geodatabase geodatabase = new(new FileGeodatabaseConnectionPath(new Uri(filePath)));

                    // Create a FeatureClassDefinition object.
                    using FeatureClassDefinition featureClassDefinition = geodatabase.GetDefinition<FeatureClassDefinition>(fileName);

                    if (featureClassDefinition != null)
                        exists = true;
                });
            }
            catch (GeodatabaseNotFoundOrOpenedException)
            {
                // Handle Exception.
                return false;
            }
            catch (GeodatabaseTableException)
            {
                // Handle Exception.
                return false;
            }

            return exists;
        }

        /// <summary>
        /// Check if a layer exists in a geodatabase.
        /// </summary>
        /// <param name="filePath">The file path to the geodatabase.</param>
        /// <param name="fileName">The table name.</param>
        /// <returns>True if the table exists; false otherwise.</returns>
        public static async Task<bool> TableExistsGDBAsync(string filePath, string fileName)
        {
            // Check there is an input file path.
            if (string.IsNullOrEmpty(filePath))
                return false;

            // Check there is an input file name.
            if (string.IsNullOrEmpty(fileName))
                return false;

            bool exists = false;

            try
            {
                await QueuedTask.Run(() =>
                {
                    // Open the file geodatabase. This will open the geodatabase if the folder exists and contains a valid geodatabase.
                    using Geodatabase geodatabase = new(new FileGeodatabaseConnectionPath(new Uri(filePath)));

                    // Create a TableDefinition object.
                    using TableDefinition tableDefinition = geodatabase.GetDefinition<TableDefinition>(fileName);

                    if (tableDefinition != null)
                        exists = true;
                });
            }
            catch (GeodatabaseNotFoundOrOpenedException)
            {
                // Handle Exception.
                return false;
            }
            catch (GeodatabaseTableException)
            {
                // Handle Exception.
                return false;
            }

            return exists;
        }

        /// <summary>
        /// Delete a feature class from a geodatabase.
        /// </summary>
        /// <param name="filePath">The file path to the geodatabase.</param>
        /// <param name="fileName">The feature class name.</param>
        /// <returns>True if the feature class was deleted
        public static async Task<bool> DeleteGeodatabaseFCAsync(string filePath, string fileName)
        {
            // Check there is an input file path.
            if (string.IsNullOrEmpty(filePath))
                return false;

            // Check there is an input file name.
            if (string.IsNullOrEmpty(fileName))
                return false;

            bool success = false;

            try
            {
                await QueuedTask.Run(() =>
                {
                    // Open the file geodatabase. This will open the geodatabase if the folder exists and contains a valid geodatabase.
                    using Geodatabase geodatabase = new(new FileGeodatabaseConnectionPath(new Uri(filePath)));

                    // Create a SchemaBuilder object
                    SchemaBuilder schemaBuilder = new(geodatabase);

                    // Create a FeatureClassDescription object.
                    using FeatureClassDefinition featureClassDefinition = geodatabase.GetDefinition<FeatureClassDefinition>(fileName);

                    // Create a FeatureClassDescription object
                    FeatureClassDescription featureClassDescription = new(featureClassDefinition);

                    // Add the deletion for the feature class to the list of DDL tasks
                    schemaBuilder.Delete(featureClassDescription);

                    // Execute the DDL
                    success = schemaBuilder.Build();
                });
            }
            catch (GeodatabaseNotFoundOrOpenedException)
            {
                // Handle Exception.
                return false;
            }
            catch (GeodatabaseTableException)
            {
                // Handle Exception.
                return false;
            }

            return success;
        }

        /// <summary>
        /// Delete a feature class from a geodatabase.
        /// </summary>
        /// <param name="geodatabase">The geodatabase to delete the feature class from.</param>
        /// <param name="featureClassName">The feature class name to delete.</param>
        /// <returns>True if the feature class was deleted
        public static async Task<bool> DeleteGeodatabaseFCAsync(Geodatabase geodatabase, string featureClassName)
        {
            // Check there is an input geodatabase.
            if (geodatabase == null)
                return false;

            // Check there is an input feature class name.
            if (string.IsNullOrEmpty(featureClassName))
                return false;

            bool success = false;

            try
            {
                await QueuedTask.Run(() =>
                {
                    // Create a SchemaBuilder object
                    SchemaBuilder schemaBuilder = new(geodatabase);

                    // Create a FeatureClassDescription object.
                    using FeatureClassDefinition featureClassDefinition = geodatabase.GetDefinition<FeatureClassDefinition>(featureClassName);

                    // Create a FeatureClassDescription object
                    FeatureClassDescription featureClassDescription = new(featureClassDefinition);

                    // Add the deletion for the feature class to the list of DDL tasks
                    schemaBuilder.Delete(featureClassDescription);

                    // Execute the DDL
                    success = schemaBuilder.Build();
                });
            }
            catch
            {
                // Handle exception.
                return false;
            }

            return success;
        }

        /// <summary>
        /// Delete a table from a geodatabase.
        /// </summary>
        /// <param name="filePath">The file path to the geodatabase to delete the table from.</param>
        /// <param name="fileName">The table name to delete.</param>
        /// <returns>True if the table was deleted</returns>
        public static async Task<bool> DeleteGeodatabaseTableAsync(string filePath, string fileName)
        {
            // Check there is an input file path.
            if (string.IsNullOrEmpty(filePath))
                return false;

            // Check there is an input file name.
            if (string.IsNullOrEmpty(fileName))
                return false;

            bool success = false;

            try
            {
                await QueuedTask.Run(() =>
                {
                    // Open the file geodatabase. This will open the geodatabase if the folder exists and contains a valid geodatabase.
                    using Geodatabase geodatabase = new(new FileGeodatabaseConnectionPath(new Uri(filePath)));

                    // Create a SchemaBuilder object
                    SchemaBuilder schemaBuilder = new(geodatabase);

                    // Create a FeatureClassDescription object.
                    using TableDefinition tableDefinition = geodatabase.GetDefinition<TableDefinition>(fileName);

                    // Create a FeatureClassDescription object
                    TableDescription tableDescription = new(tableDefinition);

                    // Add the deletion for the feature class to the list of DDL tasks
                    schemaBuilder.Delete(tableDescription);

                    // Execute the DDL
                    success = schemaBuilder.Build();
                });
            }
            catch
            {
                return false;
            }

            return success;
        }

        /// <summary>
        /// Delete a table from a geodatabase.
        /// </summary>
        /// <param name="geodatabase">The geodatabase to delete the table from.</param>
        /// <param name="tableName">The table name to delete.</param>
        /// <returns>True if the table was deleted</returns>
        public static async Task<bool> DeleteGeodatabaseTableAsync(Geodatabase geodatabase, string tableName)
        {
            // Check if the is an input geodatabase
            if (geodatabase == null)
                return false;

            // Check if there is an input table name.
            if (string.IsNullOrEmpty(tableName))
                return false;

            bool success = false;

            try
            {
                await QueuedTask.Run(() =>
                {
                    // Create a SchemaBuilder object
                    SchemaBuilder schemaBuilder = new(geodatabase);

                    // Create a FeatureClassDescription object.
                    using TableDefinition tableDefinition = geodatabase.GetDefinition<TableDefinition>(tableName);

                    // Create a FeatureClassDescription object
                    TableDescription tableDescription = new(tableDefinition);

                    // Add the deletion for the feature class to the list of DDL tasks
                    schemaBuilder.Delete(tableDescription);

                    // Execute the DDL
                    success = schemaBuilder.Build();
                });
            }
            catch
            {
                return false;
            }

            return success;
        }

        #endregion Geodatabase

        #region GeoPackage

        /// <summary>
        /// Create a new GeoPackage (.gpkg) using the Create SQLite Database geoprocessing tool.
        /// </summary>
        /// <param name="fullPath">Full path to the .gpkg file to create.</param>
        /// <returns>bool</returns>
        public static async Task<bool> CreateGeoPackageAsync(string fullPath)
        {
            // Check if there is an input full path.
            if (string.IsNullOrEmpty(fullPath))
                return false;

            // Make a value array of strings to be passed to the tool.
            // Note: The tool will create a GeoPackage when spatial_type is GEOPACKAGE.
            var parameters = Geoprocessing.MakeValueArray(fullPath, "GEOPACKAGE");

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geoprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread;

            //Geoprocessing.OpenToolDialog("management.CreateSQLiteDatabase", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.CreateSQLiteDatabase", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        #endregion GeoPackage

        #region Table

        /// <summary>
        /// Check if a feature class exists in the file path.
        /// </summary>
        /// <param name="filePath">The file path to the feature class.</param>
        /// <param name="fileName">The table name to check exists.</param>
        /// <returns>True if the table exists; false otherwise.</returns>
        public static async Task<bool> TableExistsAsync(string filePath, string fileName)
        {
            // Check there is an input file path.
            if (string.IsNullOrEmpty(filePath))
                return false;

            // Check there is an input file name.
            if (string.IsNullOrEmpty(fileName))
                return false;

            if (fileName.Substring(fileName.Length - 4, 1) == ".")
            {
                // It's a file.
                if (FileFunctions.FileExists(filePath + @"\" + fileName))
                    return true;
                else
                    return false;
            }
            else if (filePath.Substring(filePath.Length - 3, 3).Equals("sde", StringComparison.OrdinalIgnoreCase))
            {
                // It's an SDE class. Not handled (use SQL Server Functions).
                return false;
            }
            else // it is a geodatabase class.
            {
                try
                {
                    bool exists = await TableExistsGDBAsync(filePath, fileName);

                    return exists;
                }
                catch
                {
                    // GetDefinition throws an exception if the definition doesn't exist.
                    return false;
                }
            }
        }

        /// <summary>
        /// Check if a feature class exists.
        /// </summary>
        /// <param name="fullPath">The full path to the feature class.</param>
        /// <returns>True if the table exists; false otherwise.</returns>
        public static async Task<bool> TableExistsAsync(string fullPath)
        {
            // Check there is an input full path.
            if (string.IsNullOrEmpty(fullPath))
                return false;

            return await TableExistsAsync(FileFunctions.GetDirectoryName(fullPath), FileFunctions.GetFileName(fullPath));
        }

        #endregion Table

        #region Outputs

        /// <summary>
        /// Prompt the user to specify an output file in the required format.
        /// </summary>
        /// <param name="fileType">The file type.</param>
        /// <param name="initialDirectory">The initial directory.</param>
        /// <returns>The output file name, or null if the user cancelled.</returns>
        public static string GetOutputFileName(string fileType, string initialDirectory = @"C:\")
        {
            BrowseProjectFilter bf = fileType switch
            {
                "Geodatabase FC" => BrowseProjectFilter.GetFilter("esri_browseDialogFilters_geodatabaseItems_featureClasses"),
                "Geodatabase Table" => BrowseProjectFilter.GetFilter("esri_browseDialogFilters_geodatabaseItems_tables"),
                "Shapefile" => BrowseProjectFilter.GetFilter("esri_browseDialogFilters_shapefiles"),
                "CSV file (comma delimited)" => BrowseProjectFilter.GetFilter("esri_browseDialogFilters_textFiles_csv"),
                "Text file (tab delimited)" => BrowseProjectFilter.GetFilter("esri_browseDialogFilters_textFiles_txt"),
                _ => BrowseProjectFilter.GetFilter("esri_browseDialogFilters_all"),
            };

            // Display the saveItemDlg in an Open Item dialog.
            SaveItemDialog saveItemDlg = new()
            {
                Title = "Save Output As...",
                InitialLocation = initialDirectory,
                //AlwaysUseInitialLocation = true,
                //Filter = ItemFilters.Files_All,
                OverwritePrompt = false,    // This will be done later.
                BrowseFilter = bf
            };

            bool? ok = saveItemDlg.ShowDialog();

            string strOutFile = null;
            if (ok.HasValue)
                strOutFile = saveItemDlg.FilePath;

            return strOutFile; // Null if user pressed exit
        }

        #endregion Outputs

        #region CopyFeatures

        /// <summary>
        /// Copy the input feature class to the output feature class.
        /// </summary>
        /// <param name="inFeatureClass">The input feature class.</param>
        /// <param name="outFeatureClass">The output feature class.</param>
        /// <param name="addToMap">Whether to add the output to the map.</param>
        /// <returns>True if the features were copied, false if not.</returns>
        public static async Task<bool> CopyFeaturesAsync(string inFeatureClass, string outFeatureClass, bool addToMap = false)
        {
            // Check if there is an input feature class.
            if (string.IsNullOrEmpty(inFeatureClass))
                return false;

            // Check if there is an output feature class.
            if (string.IsNullOrEmpty(outFeatureClass))
                return false;

            // Make a value array of strings to be passed to the tool.
            var parameters = Geoprocessing.MakeValueArray(inFeatureClass, outFeatureClass);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geoprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; // | GPExecuteToolFlags.RefreshProjectItems;
            if (addToMap)
                executeFlags |= GPExecuteToolFlags.AddOutputsToMap;

            //Geoprocessing.OpenToolDialog("management.CopyFeatures", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.CopyFeatures", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        /// <summary>
        /// Copy the input dataset name to the output feature class.
        /// </summary>
        /// <param name="inputWorkspace">The input workspace.</param>
        /// <param name="inputDatasetName">The input dataset name.</param>
        /// <param name="outputFeatureClass">The output feature class.</param>
        /// <param name="addToMap">Whether to add the output to the map.</param>
        /// <returns>True if the features were copied, false if not.</returns>
        public static async Task<bool> CopyFeaturesAsync(string inputWorkspace, string inputDatasetName, string outputFeatureClass, bool addToMap = false)
        {
            // Check there is an input workspace.
            if (string.IsNullOrEmpty(inputWorkspace))
                return false;

            // Check there is an input dataset name.
            if (string.IsNullOrEmpty(inputDatasetName))
                return false;

            // Check there is an output feature class.
            if (string.IsNullOrEmpty(outputFeatureClass))
                return false;

            string inFeatureClass = inputWorkspace + @"\" + inputDatasetName;

            return await CopyFeaturesAsync(inFeatureClass, outputFeatureClass, addToMap);
        }

        /// <summary>
        /// Copy the input dataset to the output dataset.
        /// </summary>
        /// <param name="inputWorkspace">The input workspace.</param>
        /// <param name="inputDatasetName">The input dataset name.</param>
        /// <param name="outputWorkspace">The output workspace.</param>
        /// <param name="outputDatasetName">The output dataset name.</param>
        /// <param name="addToMap">Whether to add the output to the map.</param>
        /// <returns>True if the features were copied, false if not.</returns>
        public static async Task<bool> CopyFeaturesAsync(string inputWorkspace, string inputDatasetName, string outputWorkspace, string outputDatasetName, bool addToMap = false)
        {
            // Check there is an input workspace.
            if (string.IsNullOrEmpty(inputWorkspace))
                return false;

            // Check there is an input dataset name.
            if (string.IsNullOrEmpty(inputDatasetName))
                return false;

            // Check there is an output workspace.
            if (string.IsNullOrEmpty(outputWorkspace))
                return false;

            // Check there is an output dataset name.
            if (string.IsNullOrEmpty(outputDatasetName))
                return false;

            string inFeatureClass = inputWorkspace + @"\" + inputDatasetName;
            string outFeatureClass = outputWorkspace + @"\" + outputDatasetName;

            return await CopyFeaturesAsync(inFeatureClass, outFeatureClass, addToMap);
        }

        #endregion CopyFeatures

        #region Export Features

        /// <summary>
        /// Export the input table to the output table.
        /// </summary>
        /// <param name="inTable">The input table name.</param>
        /// <param name="outTable">The output table name.</param>
        /// <param name="addToMap">Whether to add the output to the map.</param>
        /// <returns>True if the features were exported, false if not.</returns>
        public static async Task<bool> ExportFeaturesAsync(string inTable, string outTable, bool addToMap = false)
        {
            // Check there is an input table name.
            if (string.IsNullOrEmpty(inTable))
                return false;

            // Check there is an output table name.
            if (string.IsNullOrEmpty(inTable))
                return false;

            // Make a value array of strings to be passed to the tool.
            var parameters = Geoprocessing.MakeValueArray(inTable, outTable);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geoprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; // | GPExecuteToolFlags.RefreshProjectItems;
            if (addToMap)
                executeFlags |= GPExecuteToolFlags.AddOutputsToMap;

            //Geoprocessing.OpenToolDialog("conversion.ExportTable", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("conversion.ExportTable", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        #endregion Export Features

        #region Copy Table

        /// <summary>
        /// Copy the input table to the output table.
        /// </summary>
        /// <param name="inTable">The input table name.</param>
        /// <param name="outTable">The output table name.</param>
        /// <param name="addToMap">Whether to add the output to the map.</param>
        /// <returns>True if the table was copied, false if not.</returns>
        public static async Task<bool> CopyTableAsync(string inTable, string outTable, bool addToMap = false)
        {
            // Check there is an input table name.
            if (string.IsNullOrEmpty(inTable))
                return false;

            // Check there is an output table name.
            if (string.IsNullOrEmpty(inTable))
                return false;

            // Make a value array of strings to be passed to the tool.
            var parameters = Geoprocessing.MakeValueArray(inTable, outTable);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geoprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; // | GPExecuteToolFlags.RefreshProjectItems;
            if (addToMap)
                executeFlags |= GPExecuteToolFlags.AddOutputsToMap;

            //Geoprocessing.OpenToolDialog("management.CopyRows", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.CopyRows", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        /// <summary>
        /// Copy the input dataset name to the output table.
        /// </summary>
        /// <param name="inputWorkspace">The input workspace.</param>
        /// <param name="inputDatasetName">The input dataset name.</param>
        /// <param name="outputTable">The output table name.</param>
        /// <returns>True if the table was copied, false if not.</returns>
        public static async Task<bool> CopyTableAsync(string inputWorkspace, string inputDatasetName, string outputTable)
        {
            // Check there is an input workspace.
            if (string.IsNullOrEmpty(inputWorkspace))
                return false;

            // Check there is an input dataset name.
            if (string.IsNullOrEmpty(inputDatasetName))
                return false;

            // Check there is an output feature class.
            if (string.IsNullOrEmpty(outputTable))
                return false;

            string inputTable = inputWorkspace + @"\" + inputDatasetName;

            return await CopyTableAsync(inputTable, outputTable);
        }

        /// <summary>
        /// Copy the input dataset to the output dataset.
        /// </summary>
        /// <param name="inputWorkspace">The input workspace.</param>
        /// <param name="inputDatasetName">The input dataset name.</param>
        /// <param name="outputWorkspace">The output workspace.</param>
        /// <param name="outputDatasetName">The output dataset name.</param>
        /// <returns>True if the table was copied, false if not.</returns>
        public static async Task<bool> CopyTableAsync(string inputWorkspace, string inputDatasetName, string outputWorkspace, string outputDatasetName)
        {
            // Check there is an input workspace.
            if (string.IsNullOrEmpty(inputWorkspace))
                return false;

            // Check there is an input dataset name.
            if (string.IsNullOrEmpty(inputDatasetName))
                return false;

            // Check there is an output workspace.
            if (string.IsNullOrEmpty(outputWorkspace))
                return false;

            // Check there is an output dataset name.
            if (string.IsNullOrEmpty(outputDatasetName))
                return false;

            string inputTable = inputWorkspace + @"\" + inputDatasetName;
            string outputTable = outputWorkspace + @"\" + outputDatasetName;

            return await CopyTableAsync(inputTable, outputTable);
        }

        #endregion Copy Table
    }
}
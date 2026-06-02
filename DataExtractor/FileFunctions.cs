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

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace DataTools
{
    /// <summary>
    /// This class provides basic file and folder functions.
    /// </summary>
    internal static class FileFunctions
    {
        #region Directories

        /// <summary>
        /// Check if a directory exists.
        /// </summary>
        /// <param name="filePath">The full path of the directory to check.</param>
        /// <returns>True if it exists, else false.</returns>
        public static bool DirExists(string filePath)
        {
            // Check input first.
            if (string.IsNullOrEmpty(filePath))
                return false;

            // Check if the directory exists.
            DirectoryInfo myDir = new(filePath);
            if (!myDir.Exists)
                return false;

            return true;
        }

        /// <summary>
        /// Get the name of a directory from a full path.
        /// </summary>
        /// <param name="fullPath">The full path of the file.</param>
        /// <returns>The directory name as a string.</returns>
        public static string GetDirectoryName(string fullPath)
        {
            // Check input first.
            if (string.IsNullOrEmpty(fullPath))
                return null;

            // Get the directory name.
            FileInfo fileInfo = new(fullPath);
            string dirName = fileInfo.DirectoryName;

            return dirName;
        }

        #endregion Directories

        #region Files

        /// <summary>
        /// Check if a file exists from a file path and name.
        /// </summary>
        /// <param name="filePath">The full path of the directory.</param>
        /// <param name="fileName">The name of the file.</param>
        /// <returns>True if it exists, else false.</returns>
        public static bool FileExists(string filePath, string fileName)
        {
            // Check input first.
            if (string.IsNullOrEmpty(filePath))
                return false;

            // Check input first.
            if (string.IsNullOrEmpty(fileName))
                return false;

            // If the directory exists.
            if (DirExists(filePath))
            {
                string strFileName;
                string pathEnd = filePath.Substring(filePath.Length - 1, 1);
                if (pathEnd != @"\")
                {
                    strFileName = filePath + @"\" + fileName;
                }
                else
                {
                    strFileName = filePath + fileName;
                }

                FileInfo fileInfo = new(strFileName);

                if (fileInfo.Exists) return true;
                else return false;
            }
            return false;
        }

        /// <summary>
        /// Check if a file exists from a full path.
        /// </summary>
        /// <param name="fullPath">The full path of the file.</param>
        /// <returns>True if it exists, else false.</returns>
        public static bool FileExists(string fullPath)
        {
            // Check input first.
            if (string.IsNullOrEmpty(fullPath))
                return false;

            // Check if the file exists.
            FileInfo fileInfo = new(fullPath);
            if (fileInfo.Exists)
                return true;

            return false;
        }

        /// <summary>
        /// Get the name of a file from a full path.
        /// </summary>
        /// <param name="fullPath">The full path of the file.</param>
        /// <returns>The file name as a string.</returns>
        public static string GetFileName(string fullPath)
        {
            // Check input first.
            if (string.IsNullOrEmpty(fullPath))
                return null;

            // Get the file name.
            FileInfo fileInfo = new(fullPath);
            string fileName = fileInfo.Name;

            return fileName;
        }

        /// <summary>
        /// Get a file extension from a full path.
        /// </summary>
        /// <param name="fullPath">The full path of the file.</param>
        /// <returns>The file extension as a string.</returns>
        public static string GetExtension(string fullPath)
        {
            // Check input first.
            if (string.IsNullOrEmpty(fullPath))
                return null;

            // Get the file extension.
            FileInfo fileInfo = new(fullPath);
            string aExt = fileInfo.Extension;

            return aExt;
        }

        /// <summary>
        /// Get all files in a directory.
        /// </summary>
        /// <param name="filePath">The full path of the directory.</param>
        /// <returns>The list of files as a List of strings.</returns>
        public static List<string> GetAllFilesInDirectory(string filePath)
        {
            // Check input first.
            if (string.IsNullOrEmpty(filePath))
                return null;

            List<string> myFileList = [];
            if (DirExists(filePath))
            {
                string[] fileEntries = Directory.GetFiles(filePath);
                foreach (string aFile in fileEntries)
                {
                    myFileList.Add(aFile);
                }
            }

            return myFileList;
        }

        /// <summary>
        /// Get a full file name without the extension.
        /// </summary>
        /// <param name="fullName">The full path of the file.</param>
        /// <returns>The full file name without the extension as a string.</returns>
        public static string GetFullNameWithoutExtension(string fullName)
        {
            // Check input first.
            if (string.IsNullOrEmpty(fullName))
                return null;

            // Get the directory name.
            string filePath = GetDirectoryName(fullName);

            // Get the file name without the extension.
            string fileName = Path.GetFileNameWithoutExtension(fullName);

            return filePath + @"\" + fileName;
        }

        /// <summary>
        /// Get a full file name without the extension.
        /// </summary>
        /// <param name="fileName">The full path of the file.</param>
        /// <returns>The file name without the extension as a string.</returns>
        public static string GetFileNameWithoutExtension(string fileName)
        {
            // Check input first.
            if (string.IsNullOrEmpty(fileName))
                return null;

            // Get the file name without the extension.
            fileName = Path.GetFileNameWithoutExtension(fileName);

            return fileName;
        }

        /// <summary>
        /// Delete a file.
        /// </summary>
        /// <param name="fullPath">The full path of the file.</param>
        /// <returns>True if successful, else false.</returns>
        public static bool DeleteFile(string fullPath)
        {
            // Check input first.
            if (string.IsNullOrEmpty(fullPath))
                return false;

            if (FileExists(fullPath))
            {
                try
                {
                    File.Delete(fullPath);
                    return true;
                }
                catch
                {
                    return false;
                }
            }
            else
                return true;
        }

        /// <summary>
        /// Write a new text file with optional headers.
        /// </summary>
        /// <param name="outTable">The full path of the output text file.</param>
        /// <param name="outHeader">The optional header line to write.</param>
        /// <returns>True if successful, else false.</returns>
        public static bool WriteEmptyTextFile(string outTable, string outHeader)
        {
            // Check input first.
            if (string.IsNullOrEmpty(outTable))
                return false;

            try
            {
                // Open output file.
                StreamWriter txtFile = new(outTable, false);

                // Write the headers to the file.
                if (!string.IsNullOrEmpty(outHeader))
                    txtFile.WriteLine(outHeader);

                // Close the file.
                txtFile.Close();

                // Dispose of the object.
                txtFile.Dispose();

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static bool RenameFile(string oldPath, string newPath)
        {
            // Check input first.
            if (string.IsNullOrEmpty(oldPath))
                return false;

            // Check if input exists.
            if (!FileExists(oldPath))
                return true;

            try
            {
                File.Move(oldPath, newPath);
                return true;
            }
            catch
            {
                return false;
            }
        }

        #endregion Files

        #region Logfile

        /// <summary>
        /// Write to the end of a log file.
        /// </summary>
        /// <param name="logFile">The full path of the log file.</param>
        /// <param name="logLine">The line to write to the log file.</param>
        /// <returns>True if successful, else false.</returns>
        public static bool WriteLine(string logFile, string logLine)
        {
            // Check log file name has been supplied (even though it's not used here).
            if (string.IsNullOrEmpty(logFile))
                return false;

            try
            {
                // Write to the Trace log if it is not null (the date and time will be added later).
                if (logLine != null)
                    TraceLog(logLine);
            }
            catch
            {
                return false;
            }

            return true;
        }

        private const string ToolLogPrefix = "DE|";

        /// <summary>
        /// Writes any message to the Trace log with a timestamp.
        /// </summary>
        /// <param name="message">The message to write.</param>
        private static void TraceLog(string message)
        {
            Trace.WriteLine($"{ToolLogPrefix}{DateTime.Now:G} : {message}");
        }

        #endregion Logfile
    }
}
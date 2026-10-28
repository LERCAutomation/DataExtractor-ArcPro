// The DataTools are a suite of ArcGIS Pro addins used to extract, sync
// and manage biodiversity information from ArcGIS Pro and SQL Server
// based on pre-defined or user specified criteria.
//
// Copyright © 2024-26 Andy Foy Consulting.
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
using System.Diagnostics;
using System.IO;

namespace DataExtractor
{
    internal class FilteredTraceListener(StreamWriter writer) : TextWriterTraceListener(writer)
    {

        /// <summary>
        /// Writes a line of text to the output if the message is a tool log entry.
        /// </summary>
        /// <remarks>This method filters messages and writes only those that are recognized as tool log
        /// entries, identified by the "DE|" prefix. Messages that do not meet this criterion are ignored.</remarks>
        /// <param name="message">The message to write. Only messages that begin with "DE|" are written. If null or empty, the method does
        /// nothing.</param>
        public override void WriteLine(string message)
        {
            if (string.IsNullOrEmpty(message))
                return;

            // Allow-list: only write tool log lines
            if (!message.StartsWith("DE|", StringComparison.Ordinal))
                return;

            // Remove the "DE|" prefix before writing
            string cleanMessage = message.Substring(3);
            base.WriteLine(cleanMessage);
        }
    }
}
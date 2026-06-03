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
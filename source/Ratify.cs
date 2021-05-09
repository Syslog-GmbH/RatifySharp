/*
 * Copyright 2009, 2011, 2015, 2019 P. F. Chimento
 * Copyright 2021 Syslog GmbH <https://www.syslog.de/>
 * This file is part of RatifySharp, a port of Ratify.
 * 
 * RatifySharp is free software: you can redistribute it and/or modify it under the
 * terms of the GNU Lesser General Public License as published by the Free Software
 * Foundation, either version 3 of the License, or (at your option) any later version.
 * 
 * RatifySharp is distributed in the hope that it will be useful, but WITHOUT ANY
 * WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A
 * PARTICULAR PURPOSE. See the GNU Lesser General Public License for more details.
 * You should have received a copy of the GNU Lesser General Public License along
 * with RatifySharp. If not, see <https://www.gnu.org/licenses/>.
 */

using Gdk;
using Gtk;

using System;
using System.Text;

namespace Ratify
{
    public static class Ratify
    {
        // FIXME: Not working with GtkSharp?
        /// <summary>
        /// Registers the RTF text serialization format with <paramref name="buffer"/>. This allows the
        /// contents of <paramref name="buffer"/> to be exported to Rich Text Format (MIME type text/rtf).
        /// </summary>
        /// <param name="buffer">Text buffer</param>
        /// <returns>Gdk.Atom representing the serialization format, to be passed to gtk_text_buffer_serialize()</returns>
        public static Atom RegisterSerializeFormat(TextBuffer buffer)
        {
            if (buffer == null)
                return new Atom(IntPtr.Zero);

            return buffer.RegisterSerializeFormat("text/rtf", RtfSerialize.SerializeTarget);
        }

        // FIXME: Not working with GtkSharp?
        /// <summary>
        /// Registers the RTF text serialization format with <paramref name="buffer"/>. This allows the
        /// contents of <paramref name="buffer"/> to be exported to Rich Text Format (MIME type text/rtf).
        /// </summary>
        /// <param name="buffer">Text buffer</param>
        /// <returns>Gdk.Atom representing the serialization format, to be passed to gtk_text_buffer_serialize()</returns>
        public static Atom RegisterDeserializeFormat(TextBuffer buffer)
        {
            if (buffer == null)
                return new Atom(IntPtr.Zero);

            var format = buffer.RegisterDeserializeFormat("text/rtf", RtfDeserialize.DeserializeTarget);
            buffer.DeserializeSetCanCreateTags(format, true);
            return format;
        }

        /// <summary>
        /// Deserializes the contents of <paramref name="str"/> to <paramref name="buffer"/>.
        /// See rtf_text_buffer_import_file() for details.
        /// </summary>
        /// <param name="buffer">Text buffer into which to import text</param>
        /// <param name="str">String containing an RTF document</param>
        /// <returns>Success or failure</returns>
        /// <remarks>
        /// If <paramref name="str"/> contains references to external files, such as images, then
        /// these will be resolved relative to the current working directory.
        /// That's usually not what you want.
        /// If you want to load images, then you must change the current working
        /// directory appropriately before calling this function, or embed the images
        /// in the RTF code, or use rtf_text_buffer_import() or rtf_text_buffer_import_file().
        /// </remarks>
        public static bool Deserialize(TextBuffer buffer, string str)
        {
            if (buffer == null || str == null)
                return false;

            buffer.Clear();

            return RtfDeserialize.Deserialize(buffer, buffer.StartIter, str);
        }

        /// <summary>
        /// Serializes the contents of <paramref name="buffer"/> to a string in RTF format.
        /// See rtf_text_buffer_export() for details.
        /// </summary>
        /// <param name="buffer">Text buffer to export</param>
        /// <returns>String containing RTF text or null</returns>
        public static string Serialize(TextBuffer buffer)
        {
            if (buffer == null)
                return null;

            buffer.GetBounds(out TextIter start, out TextIter end);

            return RtfSerialize.Serialize(buffer, start, end);
        }

        /// <summary>
        /// Deserializes the contents of <paramref name="str"/> to <paramref name="buffer"/>.
        /// See rtf_text_buffer_import_file() for details.
        /// </summary>
        /// <param name="buffer">Text buffer into which to import text</param>
        /// <param name="format">Format Atom</param>
        /// <param name="str">String</param>
        /// <returns>Success or failure</returns>
        public static bool DeserializeFormat(TextBuffer buffer, Atom format, string str)
        {
            if (buffer == null || str == null)
                return false;

            buffer.Clear();

            var iter = buffer.StartIter;
            var data = Encoding.Default.GetBytes(str);

            return buffer.Deserialize(buffer, format, ref iter, data, (ulong)data.LongLength);
        }

        /// <summary>
        /// Serializes the contents of <paramref name="buffer"/> to a string in RTF format.
        /// See rtf_text_buffer_export() for details.
        /// </summary>
        /// <param name="buffer">Text buffer to export</param>
        /// <param name="format">Format Atom</param>
        /// <returns>String containing RTF text or null</returns>
        public static string SerializeFormat(TextBuffer buffer, Atom format)
        {
            if (buffer == null)
                return null;

            buffer.GetBounds(out TextIter start, out TextIter end);

            var data = buffer.Serialize(buffer, format, start, end);
            if (data == null || data.Length == 0)
                return null;

            var str = Encoding.Default.GetString(data);
            return str;
        }
    }
}

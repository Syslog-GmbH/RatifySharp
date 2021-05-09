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

using Pango;

using System;

namespace Ratify
{
    public abstract class RtfState : IDisposable, ICloneable
    {
        public abstract object Clone();
        public abstract void Dispose();
    }

    public class RtfStateAttributes : RtfState
    {
        /// <summary>
        /// Index into style sheet
        /// </summary>
        public int Style = -1;

        /// <summary>
        /// GtkJustification value or -1 if unset
        /// </summary>
        public int Justification = -1;

        /// <summary>
        /// GtkTextDirection value or -1 if unset
        /// </summary>
        public int ParDirection = -1;

        public int SpaceBefore;
        public int SpaceAfter;
        public bool IgnoreSpaceBefore;
        public bool IgnoreSpaceAfter;
        public TabArray Tabs;
        public int LeftMargin;
        public int RightMargin;
        public int Indent;
        public int Leading;

        /// <summary>
        /// Index into the color table
        /// </summary>
        public int Foreground = -1;

        /// <summary>
        /// Index into the color table
        /// </summary>
        public int Background = -1;

        /// <summary>
        /// Index into the color table
        /// </summary>
        public int Highlight = -1;

        /// <summary>
        /// Index into the font table
        /// </summary>
        public int Font = -1;

        public double Size;
        public bool Italic;
        public bool Bold;
        public bool SmallCaps;
        public bool Strikethrough;
        public bool Subscript;
        public bool Superscript;
        public bool Invisible;

        /// <summary>
        /// PangoUnderline value or -1 if unset
        /// </summary>
        public int Underline = -1;

        /// <summary>
        /// GtkTextDirection value or -1 if unset
        /// </summary>
        public int CharDirection = -1;

        public int Language = 1024;
        public int Rise;
        public int Scale = 100;

        /// <summary>
        /// Number of characters to skip after \u
        /// </summary>
        public int UnicodeSkip = 1;

        /// <summary>
        /// Skip characters within \upr but not \*ud
        /// </summary>
        public bool UnicodeIgnore;

        public void SetDefaultCharacterAttributes()
        {
            Style = -1;
            Background = -1;
            Foreground = -1;
            Highlight = -1;
            Font = -1;
            Size = 0.0;
            Italic = false;
            Bold = false;
            SmallCaps = false;
            Strikethrough = false;
            Subscript = false;
            Superscript = false;
            Underline = -1;
            CharDirection = -1;
            Language = 1024;
            Rise = 0;
            Scale = 100;
        }

        public void SetDefaultParagraphAttributes()
        {
            Justification = -1;
            ParDirection = -1;
            SpaceBefore = 0;
            SpaceAfter = 0;
            IgnoreSpaceBefore = false;
            IgnoreSpaceAfter = false;
            Tabs?.Dispose();
            Tabs = null;
            LeftMargin = 0;
            RightMargin = 0;
            Indent = 0;
            Leading = 0;
        }

        public override object Clone()
        {
            var clone = MemberwiseClone() as RtfStateAttributes;

            if (Tabs != null)
                clone.Tabs = Tabs.Copy();

            return clone;
        }

        public override void Dispose()
        {
            Tabs?.Dispose();
            Tabs = null;
        }
    }
}

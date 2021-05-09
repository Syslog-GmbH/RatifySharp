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

using Gtk;

using System;
using System.Collections.Generic;
using System.Linq;

namespace Ratify
{
    public static class RtfFontTable
    {
        public enum FontFamily
        {
            FONT_FAMILY_NIL,
            FONT_FAMILY_ROMAN,
            FONT_FAMILY_SWISS,
            FONT_FAMILY_MODERN,
            FONT_FAMILY_SCRIPT,
            FONT_FAMILY_DECORATIVE,
            FONT_FAMILY_TECH,
            FONT_FAMILY_BIDI
        }

        public class RtfFontTableState : RtfState
        {
            public int Index;
            public int Codepage = -1;
            public FontFamily Family;
            public string Name = string.Empty;

            public override object Clone()
            {
                return MemberwiseClone();
            }

            public override void Dispose()
            {
            }
        }

        /// <summary>
        /// Convert "font charset" character encoding to "codepage" character encoding
        /// </summary>
        public static int CharsetToCodepage(int charset)
        {
            switch (charset)
            {
                case 0: return 1252;  /* "ANSI" */
                case 1: return -1;    /* default */
                case 2: return -1;    /* Symbol; only works in Symbol font of course */
                case 77: return 10000; /* Mac Roman */
                case 78: return 10001; /* Mac Shift JIS */
                case 79: return 10003; /* Mac Hangul */
                case 80: return 10008; /* Mac GB2312 */
                case 81: return 10002; /* Mac Big5 */
                case 83: return 10005; /* Mac Hebrew */
                case 84: return 10004; /* Mac Arabic */
                case 85: return 10006; /* Mac Greek */
                case 86: return 10081; /* Mac Turkish */
                case 87: return 10021; /* Mac Thai */
                case 88: return 10029; /* Mac East Europe */
                case 89: return 10007; /* Mac Cyrillic */
                case 128: return 943;   /* ShiftJIS */
                case 129: return 949;   /* Hangul */
                case 130: return 1361;  /* Johab */
                case 134: return 936;   /* GB2312 */
                case 136: return 950;   /* Chinese Big5 */
                case 161: return 1253;  /* Greek */
                case 162: return 1254;  /* Turkish */
                case 163: return 1258;  /* Vietnamese */
                case 177: return 1255;  /* Hebrew */
                case 178: return 1256;  /* Arabic */
                case 181: return 862;   /* Hebrew user */
                case 186: return 1257;  /* Baltic */
                case 204: return 1251;  /* Russian */
                case 222: return 874;   /* Thai */
                case 238: return 1250;  /* Eastern European */
                case 254: return 437;   /* PC 437 */
                case 255: return 850;   /* OEM */

                case 82:  /* Mac Johab (old) */
                case 179: /* Arabic traditional (old) */
                case 180: /* Arabic user (old) */
                    //g_warning(_("Character set %d not supported"), charset);
                    return -1;
                default:
                    //g_warning(_("Unknown character set %d"), charset);
                    return -1;
            }
        }

        public static bool ft_f(RtfDeserialize.ParserContext ctx, RtfState state, int index)
        {
            if (state is RtfFontTableState fontState)
                fontState.Index = index;

            return true;
        }

        public static bool ft_fbidi(RtfDeserialize.ParserContext ctx, RtfState state, int par)
        {
            if (state is RtfFontTableState fontState)
                fontState.Family = FontFamily.FONT_FAMILY_BIDI;

            return true;
        }

        public static bool ft_fdecor(RtfDeserialize.ParserContext ctx, RtfState state, int par)
        {
            if (state is RtfFontTableState fontState)
                fontState.Family = FontFamily.FONT_FAMILY_DECORATIVE;

            return true;
        }

        public static bool ft_fmodern(RtfDeserialize.ParserContext ctx, RtfState state, int par)
        {
            if (state is RtfFontTableState fontState)
                fontState.Family = FontFamily.FONT_FAMILY_MODERN;

            return true;
        }

        public static bool ft_fnil(RtfDeserialize.ParserContext ctx, RtfState state, int par)
        {
            if (state is RtfFontTableState fontState)
                fontState.Family = FontFamily.FONT_FAMILY_NIL;

            return true;
        }

        public static bool ft_froman(RtfDeserialize.ParserContext ctx, RtfState state, int par)
        {
            if (state is RtfFontTableState fontState)
                fontState.Family = FontFamily.FONT_FAMILY_ROMAN;

            return true;
        }

        public static bool ft_fscript(RtfDeserialize.ParserContext ctx, RtfState state, int par)
        {
            if (state is RtfFontTableState fontState)
                fontState.Family = FontFamily.FONT_FAMILY_SCRIPT;

            return true;
        }

        public static bool ft_fswiss(RtfDeserialize.ParserContext ctx, RtfState state, int par)
        {
            if (state is RtfFontTableState fontState)
                fontState.Family = FontFamily.FONT_FAMILY_SWISS;

            return true;
        }

        public static bool ft_ftech(RtfDeserialize.ParserContext ctx, RtfState state, int par)
        {
            if (state is RtfFontTableState fontState)
                fontState.Family = FontFamily.FONT_FAMILY_TECH;

            return true;
        }

        public static bool ft_fcharset(RtfDeserialize.ParserContext ctx, RtfState state, int charset)
        {
            if (state is RtfFontTableState fontState)
                fontState.Codepage = CharsetToCodepage(charset);

            return true;
        }

        public static readonly List<RtfDeserialize.ControlWord> _fontTableWordTable = new List<RtfDeserialize.ControlWord>()
        {
            new RtfDeserialize.ControlWord("f", RtfDeserialize.ControlWordType.RequiredParameter, true, ft_f),
            new RtfDeserialize.ControlWord("fbidi", RtfDeserialize.ControlWordType.NoParameter, true, ft_fbidi),
            new RtfDeserialize.ControlWord("fcharset", RtfDeserialize.ControlWordType.RequiredParameter, true, ft_fcharset),
            new RtfDeserialize.ControlWord("fdecor", RtfDeserialize.ControlWordType.NoParameter, true, ft_fdecor),
            new RtfDeserialize.ControlWord("fmodern", RtfDeserialize.ControlWordType.NoParameter, true, ft_fmodern),
            new RtfDeserialize.ControlWord("fnil", RtfDeserialize.ControlWordType.NoParameter, true, ft_fnil),
            new RtfDeserialize.ControlWord("froman", RtfDeserialize.ControlWordType.NoParameter, true, ft_froman),
            new RtfDeserialize.ControlWord("fscript", RtfDeserialize.ControlWordType.NoParameter, true, ft_fscript),
            new RtfDeserialize.ControlWord("fswiss", RtfDeserialize.ControlWordType.NoParameter, true, ft_fswiss),
            new RtfDeserialize.ControlWord("ftech", RtfDeserialize.ControlWordType.NoParameter, true, ft_ftech),
        };
        
        public static List<RtfDeserialize.ControlWord> FontTableWordTable
        {
            get
            {
                var fontTableWordTable = new List<RtfDeserialize.ControlWord>();
                fontTableWordTable.AddRange(RtfDocument.SpecialCharacterControlWords);
                fontTableWordTable.AddRange(_fontTableWordTable);
                return fontTableWordTable;
            }
        }

        public class FontTableDestinationInfo : RtfDeserialize.DestinationInfo
        {
            private readonly string[] _fontSuggestions = {
                    "Sans", // Default font for \fnil
                    "Serif", // \froman
                    "Sans", // \fswiss
                    "Monospace", // \fmodern
                    "Script", // \fscript
                    null, // \fdecor
                    null, // \ftech
                    null // \fbidi
                };

            public override List<RtfDeserialize.ControlWord> ControlWords => FontTableWordTable;

            public override Type StateType => typeof(RtfFontTableState);

            public override void Cleanup(RtfDeserialize.ParserContext ctx)
            {
            }

            /// <summary>
            /// Process plain text in the font table (font names separated by semicolons)
            /// </summary>
            public override void Flush(RtfDeserialize.ParserContext ctx)
            {
                if (ctx == null)
                    throw new ArgumentNullException("ctx");

                var state = ctx.GetState() as RtfFontTableState;
                string name = ctx.Text;

                int semicolonPos = name.IndexOf(';');
                if (semicolonPos == -1)
                {
                    string newName = state.Name + name;
                    state.Name = newName;
                    ctx.Text = string.Empty;
                    return;
                }

                if (semicolonPos < name.Length - 1)
                    ctx.Text = name.Substring(semicolonPos + 1); // Leave the text after the semicolon in the buffer
                else
                    ctx.Text = string.Empty;

                name = name.Substring(0, semicolonPos);

                var fontProp = new RtfDeserialize.FontProperties();
                fontProp.Index = state.Index;
                fontProp.Codepage = state.Codepage;
                fontProp.FontName = state.Name + name;

                ctx.FontTable = ctx.FontTable.Prepend(fontProp).ToList();

                // Add the tag to the buffer right now instead of when the font is used,
                // since any font might be declared the default font; remove any previous font
                // with this font table index first
                var tagName = $"ratifysharp-font-{state.Index}";
                var tag = ctx.Tags.Lookup(tagName);
                if (tag != null)
                {
                    ctx.Tags.Remove(tag);
                }

                tag = new TextTag(tagName);

                int family = (int)state.Family;
                string fontString = null;

                if (fontProp.FontName != null && _fontSuggestions[family] != null)
                    fontString = fontProp.FontName + "," + _fontSuggestions[family];
                else if (fontProp.FontName != null)
                    fontString = fontProp.FontName;
                else if (_fontSuggestions[family] != null)
                    fontString = _fontSuggestions[family];

                if (fontString != null)
                {
                    tag.Family = fontString;
                    tag.FamilySet = true;
                }

                ctx.Tags.Add(tag);

                state.Index = 0;
                state.Family = FontFamily.FONT_FAMILY_NIL;
                state.Codepage = -1;
                state.Name = string.Empty;
            }

            /// <summary>
            /// Assume that text in the \fonttbl destination is in the encoding specified by that entry's \fcharset
            /// </summary>
            public override int GetCodepage(RtfDeserialize.ParserContext ctx)
            {
                var state = ctx.GetState() as RtfFontTableState;
                if (state != null)
                {
                    return state.Codepage;
                }

                return -1;
            }
        }
    }
}

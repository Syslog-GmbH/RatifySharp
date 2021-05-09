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

using System;
using System.Collections.Generic;

namespace Ratify
{
    public static class RtfFootnote
    {
        public static List<RtfDeserialize.ControlWord> FootnoteWordTable => RtfDocument.DocumentTextControlWords;

        public class FootnoteDestinationInfo : RtfDeserialize.DestinationInfo
        {
            public override List<RtfDeserialize.ControlWord> ControlWords => FootnoteWordTable;

            public override Type StateType => typeof(RtfStateAttributes);

            public override void Cleanup(RtfDeserialize.ParserContext ctx)
            {
            }

            /// <summary>
            /// This function is mostly the same as document_text(),
            /// but adds the text to the end of the textbuffer.
            /// </summary>
            public override void Flush(RtfDeserialize.ParserContext ctx)
            {
                if (ctx == null)
                    throw new ArgumentNullException("ctx");

                string text = ctx.Text;
                int length = text.Length - 1;
                if (length < 0)
                    return;

                if (ctx.GroupNestingLevel == 0 && text[length] == '\n')
                    text.Remove(length);
                
                var attr = ctx.GetState() as RtfStateAttributes;
                if (attr != null)
                {
                    var end = ctx.TextBuffer.EndIter;
                    var placeholder = ctx.TextBuffer.CreateMark(null, end, true);
                    ctx.TextBuffer.Insert(ref end, text);

                    var start = ctx.TextBuffer.GetIterAtMark(placeholder);
                    end = ctx.TextBuffer.EndIter;

                    ctx.ApplyAttributes(attr, start, end);

                    ctx.TextBuffer.DeleteMark(placeholder);

                    // Move the regular document endmark back to the startmark so that
                    // subsequent document text is inserted before the footnotes
                    start = ctx.TextBuffer.GetIterAtMark(ctx.StartMark);
                    ctx.TextBuffer.MoveMark(ctx.EndMark, start);
                }

                ctx.Text = string.Empty;
            }

            /// <summary>
            /// Return the codepage associated with the current font.
            /// </summary>
            public override int GetCodepage(RtfDeserialize.ParserContext ctx)
            {
                var attr = ctx.GetState() as RtfStateAttributes;
                if (attr != null)
                {
                    if (attr.Font != -1)
                    {
                        var fontProp = ctx.GetFontProperties(attr.Font);
                        if (fontProp != null)
                        {
                            return fontProp.Codepage;
                        }
                    }
                }

                return -1;
            }
        }
    }
}

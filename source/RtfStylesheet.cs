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
using Pango;

using System;
using System.Collections.Generic;

namespace Ratify
{
    public static class RtfStylesheet
    {
        public enum StyleType
        {
            STYLE_PARAGRAPH,
            STYLE_CHARACTER,
            STYLE_SECTION,
            STYLE_TABLE
        }

        public class RtfStylesheetState : RtfState
        {
            public RtfStateAttributes Attributes = new RtfStateAttributes();
            public int Index = 0;
            public StyleType Style = StyleType.STYLE_PARAGRAPH;

            public override object Clone()
            {
                return MemberwiseClone();
            }

            public override void Dispose()
            {
            }
        }
        

        public static bool sty_cs(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStylesheetState styleState)
            {
                styleState.Index = param;
                styleState.Style = StyleType.STYLE_CHARACTER;
            }

            return true;
        }

        public static bool sty_ds(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStylesheetState styleState)
            {
                styleState.Index = param;
                styleState.Style = StyleType.STYLE_SECTION;
            }

            return true;
        }

        public static bool sty_s(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStylesheetState styleState)
            {
                styleState.Index = param;
                styleState.Style = StyleType.STYLE_PARAGRAPH;
            }

            return true;
        }

        public static bool sty_ts(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStylesheetState styleState)
            {
                styleState.Index = param;
                styleState.Style = StyleType.STYLE_TABLE;
            }

            return true;
        }

        public static readonly List<RtfDeserialize.ControlWord> _stylesheetWordTable = new List<RtfDeserialize.ControlWord>()
        {
            new RtfDeserialize.ControlWord("*cs", RtfDeserialize.ControlWordType.RequiredParameter, true, sty_cs),
            new RtfDeserialize.ControlWord("*ds", RtfDeserialize.ControlWordType.RequiredParameter, true, sty_ds),
            new RtfDeserialize.ControlWord("s", RtfDeserialize.ControlWordType.OptionalParameter, true, sty_s),
            new RtfDeserialize.ControlWord("*ts", RtfDeserialize.ControlWordType.RequiredParameter, true, sty_ts)
        };

        public static List<RtfDeserialize.ControlWord> StylesheetWordTable
        {
            get
            {
                var stylesheetWordTable = new List<RtfDeserialize.ControlWord>();
                stylesheetWordTable.AddRange(RtfDocument.FormattedTextControlWords);
                stylesheetWordTable.AddRange(_stylesheetWordTable);
                return stylesheetWordTable;
            }
        }

        public class StylesheetDestinationInfo : RtfDeserialize.DestinationInfo
        {
            public override List<RtfDeserialize.ControlWord> ControlWords => StylesheetWordTable;

            public override Type StateType => typeof(RtfStylesheetState);

            public override void Cleanup(RtfDeserialize.ParserContext ctx)
            {
            }

            /// <summary>
            /// Add a style tag to the GtkTextBuffer's tag table with all the attributes of
            /// the current style
            /// </summary>
            public override void Flush(RtfDeserialize.ParserContext ctx)
            {
                if (ctx == null)
                    throw new ArgumentNullException("ctx");

                var state = ctx.GetState() as RtfStylesheetState;
                if (state != null)
                {
                    int semicolonPos = ctx.Text.IndexOf(';');
                    if (semicolonPos == -1 || semicolonPos == ctx.Text.Length - 1)
                    {
                        ctx.Text = string.Empty;
                        return;
                    }

                    ctx.Text = ctx.Text.Substring(semicolonPos + 1); // Leave the text after the semicolon in the buffer

                    var tagName = $"ratifysharp-style-{state.Index}";
                    var tag = ctx.Tags.Lookup(tagName);
                    if (tag != null)
                    {
                        ctx.Tags.Remove(tag);
                    }

                    tag = new TextTag(tagName);

                    // Add each paragraph attribute to the tag
                    if (state.Attributes.Justification != -1)
                    {
                        tag.Justification = (Justification)state.Attributes.Justification;
                        tag.JustificationSet = true;
                    }

                    if (state.Attributes.ParDirection != -1)
                    {
                        tag.Direction = (TextDirection)state.Attributes.ParDirection;
                    }

                    if (state.Attributes.SpaceBefore != 0 && !state.Attributes.IgnoreSpaceBefore)
                    {
                        tag.PixelsAboveLines = Units.ToPixels(RtfDeserialize.TwipsToPango(state.Attributes.SpaceBefore));
                        tag.PixelsAboveLinesSet = true;
                    }

                    if (state.Attributes.SpaceAfter != 0 && !state.Attributes.IgnoreSpaceAfter)
                    {
                        tag.PixelsBelowLines = Units.ToPixels(RtfDeserialize.TwipsToPango(state.Attributes.SpaceAfter));
                        tag.PixelsBelowLinesSet = true;
                    }

                    if (state.Attributes.Tabs != null)
                    {
                        tag.Tabs = state.Attributes.Tabs;
                        tag.TabsSet = true;
                    }

                    if (state.Attributes.LeftMargin != 0)
                    {
                        tag.LeftMargin = Units.ToPixels(RtfDeserialize.TwipsToPango(state.Attributes.LeftMargin));
                        tag.LeftMarginSet = true;
                    }

                    if (state.Attributes.RightMargin != 0)
                    {
                        tag.RightMargin = Units.ToPixels(RtfDeserialize.TwipsToPango(state.Attributes.RightMargin));
                        tag.RightMarginSet = true;
                    }

                    if (state.Attributes.Indent != 0)
                    {
                        tag.Indent = Units.ToPixels(RtfDeserialize.TwipsToPango(state.Attributes.Indent));
                        tag.IndentSet = true;
                    }

                    if (state.Attributes.Scale != 100)
                    {
                        tag.Scale = state.Attributes.Scale / 100.0;
                        tag.ScaleSet = true;
                    }

                    // Add each character attribute to the tag
                    if (state.Attributes.Foreground != -1)
                    {
                        var color = ctx.ColorTable[state.Attributes.Foreground];

                        // color must exist, because that was already checked when executing the \cf command
                        tag.Foreground = color;
                        tag.ForegroundSet = true;
                    }

                    if (state.Attributes.Background != -1)
                    {
                        var color = ctx.ColorTable[state.Attributes.Background];

                        // color must exist, because that was already checked when executing the \cf command
                        tag.Background = color;
                        tag.BackgroundSet = true;
                    }

                    if (state.Attributes.Highlight != -1)
                    {
                        var color = ctx.ColorTable[state.Attributes.Highlight];

                        // color must exist, because that was already checked when executing the \cf command
                        tag.ParagraphBackground = color;
                        tag.SetProperty("paragraph-background-set", new GLib.Value(true));
                    }

                    if (state.Attributes.Font != -1)
                    {
                        var tagNameFont = $"ratifysharp-font-{state.Attributes.Font}";
                        var fontTag = ctx.Tags.Lookup(tagNameFont);
                        if (fontTag != null)
                        {
                            var fontDesc = fontTag.FontDesc;
                            if (fontDesc != null && !string.IsNullOrWhiteSpace(fontDesc.Family))
                            {
                                tag.Family = fontDesc.Family;
                                tag.FamilySet = true;
                            }
                        }
                    }

                    if (state.Attributes.Size != 0.0)
                    {
                        tag.Size = RtfDeserialize.PointsToPango(state.Attributes.Size);
                        tag.SizePoints = state.Attributes.Size;
                        tag.SizeSet = true;
                    }

                    if (state.Attributes.Italic)
                    {
                        tag.Style = Pango.Style.Italic;
                        tag.StyleSet = true;
                    }

                    if (state.Attributes.Bold)
                    {
                        tag.Weight = Weight.Bold;
                        tag.WeightSet = true;
                    }

                    if (state.Attributes.SmallCaps)
                    {
                        tag.Variant = Pango.Variant.SmallCaps;
                        tag.VariantSet = true;
                    }

                    if (state.Attributes.Strikethrough)
                    {
                        tag.Strikethrough = true;
                        tag.StrikethroughSet = true;
                    }

                    if (state.Attributes.Subscript)
                    {
                        tag.Rise = RtfDeserialize.PointsToPango(-6);
                        tag.RiseSet = true;
                        tag.Scale = Pango.Scale.XSmall;
                        tag.ScaleSet = true;
                    }

                    if (state.Attributes.Superscript)
                    {
                        tag.Rise = RtfDeserialize.PointsToPango(6);
                        tag.RiseSet = true;
                        tag.Scale = Pango.Scale.XSmall;
                        tag.ScaleSet = true;
                    }

                    if (state.Attributes.Invisible)
                    {
                        tag.Invisible = true;
                        tag.InvisibleSet = true;
                    }

                    if (state.Attributes.Underline != -1)
                    {
                        tag.Underline = (Underline)state.Attributes.Underline;
                        tag.UnderlineSet = true;
                    }

                    if (state.Attributes.CharDirection != -1)
                    {
                        tag.Direction = (TextDirection)state.Attributes.CharDirection;
                    }

                    if (state.Attributes.Rise != 0)
                    {
                        tag.Rise = RtfDeserialize.HalfPointsToPango(state.Attributes.Rise);
                        tag.RiseSet = true;
                    }

                    ctx.Tags.Add(tag);

                    state.Index = 0;
                    state.Style = StyleType.STYLE_PARAGRAPH;
                    state.Attributes.SetDefaultParagraphAttributes();
                    state.Attributes.SetDefaultCharacterAttributes();
                }
            }
            
            public override int GetCodepage(RtfDeserialize.ParserContext ctx)
            {
                return 0;
            }
        }
    }
}

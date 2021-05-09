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
    public static class RtfDocument
    {
        public static bool doc_u(RtfDeserialize.ParserContext ctx, RtfState state, int ch)
        {
            if (state is RtfStateAttributes attr)
            {
                int code = 0;
                if (ch < 0)
                    code = ch + 65536;
                else
                    code = ch;

                char utf8 = Convert.ToChar(code);
                ctx.Text += utf8;

                for (int foo = 0; foo < attr.UnicodeSkip; foo++)
                    if (!ctx.SkipCharacterOrControlWord())
                        return false;
            }

            return true;
        }

        public static bool doc_uc(RtfDeserialize.ParserContext ctx, RtfState state, int skip)
        {
            if (state is RtfStateAttributes attr)
                attr.UnicodeSkip = skip;

            return true;
        }

        public static readonly List<RtfDeserialize.ControlWord> SpecialCharacterControlWords = new List<RtfDeserialize.ControlWord>()
        {
            new RtfDeserialize.ControlWord("\n", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\n"),
            new RtfDeserialize.ControlWord("\r", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\n"),
            new RtfDeserialize.ControlWord("-", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\xC2\xAD"), // U+00AD Soft hyphen
            new RtfDeserialize.ControlWord("\\", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\\"),
            new RtfDeserialize.ControlWord("_", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\xE2\x80\x91"), // U+2011 NBhyphen
            new RtfDeserialize.ControlWord("{", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "{"),
            new RtfDeserialize.ControlWord("}", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "}"),
            new RtfDeserialize.ControlWord("~", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\xC2\xA0"), // U+00A0 NBSP
            new RtfDeserialize.ControlWord("bullet", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\xE2\x80\xA2"), // U+2022 Bullet
            new RtfDeserialize.ControlWord("emdash", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\xE2\x80\x94"), // U+2014 em dash
            new RtfDeserialize.ControlWord("emspace", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\xE2\x80\x83"), // U+2003 em space
            new RtfDeserialize.ControlWord("endash", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\xE2\x80\x93"), // U+2013 en dash
            new RtfDeserialize.ControlWord("enspace", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\xE2\x80\x82"), // U+2002 en space
            new RtfDeserialize.ControlWord("line", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\xE2\x80\xA8"), // U+2028 Line separator
            new RtfDeserialize.ControlWord("ldblquote", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\xE2\x80\x9C"), // U+201C Left double quote
            new RtfDeserialize.ControlWord("lquote", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\xE2\x80\x98"), // U+2018 Left single quote
            new RtfDeserialize.ControlWord("ltrmark", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\xE2\x80\x8E"), // U+200E Left-to-right mark
            new RtfDeserialize.ControlWord("par", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\n"),
            new RtfDeserialize.ControlWord("qmspace", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\xE2\x80\x85"), // U+2005 4 per em space
            new RtfDeserialize.ControlWord("rdblquote", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\xE2\x80\x9D"), // U+201D Right double quote
            new RtfDeserialize.ControlWord("rquote", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\xE2\x80\x99"), // U+2019 Right single quote
            new RtfDeserialize.ControlWord("rtlmark", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\xE2\x80\x8F"), // U+200F Right-to-left mark
            new RtfDeserialize.ControlWord("tab", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\t"),
            new RtfDeserialize.ControlWord("u", RtfDeserialize.ControlWordType.RequiredParameter, false, doc_u),
            new RtfDeserialize.ControlWord("uc", RtfDeserialize.ControlWordType.RequiredParameter, false, doc_uc),
            new RtfDeserialize.ControlWord("zwbo", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\xE2\x80\x8B"), // U+200B zero width space
            new RtfDeserialize.ControlWord("zwj", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\xE2\x80\x8D"), // U+200D zero width joiner
            new RtfDeserialize.ControlWord("zwnj", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\xE2\x80\x8C") // U+200C zero width non joiner
        };

        public static bool doc_b(RtfDeserialize.ParserContext ctx, RtfState state, int par)
        {
            if (state is RtfStateAttributes attr)
            {
                string tagName = "ratifysharp-bold";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Weight = Weight.Bold;
                    tag.WeightSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.Bold = par != 0;
            }

            return true;
        }

        public static bool doc_cb(RtfDeserialize.ParserContext ctx, RtfState state, int par)
        {
            if (ctx.ColorTable.Count > par)
            {
                //g_set_error(error, RTF_ERROR, RTF_ERROR_UNDEFINED_COLOR, _("Color '%i' undefined"), param);
                return false;
            }

            string color = ctx.ColorTable[par];

            if (state is RtfStateAttributes attr)
            {
                string tagName = $"ratifysharp-background-{par}";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Background = color;
                    tag.BackgroundSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.Background = par;
            }

            return true;
        }

        public static bool doc_cf(RtfDeserialize.ParserContext ctx, RtfState state, int par)
        {
            if (ctx.ColorTable.Count > par)
            {
                //g_set_error(error, RTF_ERROR, RTF_ERROR_UNDEFINED_COLOR, _("Color '%i' undefined"), param);
                return false;
            }

            string color = ctx.ColorTable[par];

            if (state is RtfStateAttributes attr)
            {
                string tagName = $"ratifysharp-foreground-{par}";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Foreground = color;
                    tag.ForegroundSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.Foreground = par;
            }

            return true;
        }

        public static bool doc_charscalex(RtfDeserialize.ParserContext ctx, RtfState state, int scale)
        {
            if (scale <= 0)
            {
                //g_set_error(error, RTF_ERROR, RTF_ERROR_BAD_FONT_SIZE, _("\\charscalex%d is invalid, negative or zero scales not allowed"), scale);
                return false;
            }

            if (state is RtfStateAttributes attr)
            {
                string tagName = $"ratifysharp-scale-{scale}";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Scale = scale / 100.0;
                    tag.ScaleSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.Scale = scale;
            }

            return true;
        }

        public static bool doc_dn(RtfDeserialize.ParserContext ctx, RtfState state, int halfpoints)
        {
            if (state is RtfStateAttributes attr)
            {
                if (halfpoints != 0)
                {
                    string tagName = $"ratifysharp-down-{halfpoints}";
                    if (ctx.Tags.Lookup(tagName) == null)
                    {
                        var tag = new TextTag(tagName);
                        tag.Rise = RtfDeserialize.HalfPointsToPango(-halfpoints);
                        tag.RiseSet = true;
                        ctx.Tags.Add(tag);
                    }
                }

                attr.Rise = -halfpoints;
            }

            return true;
        }

        public static bool doc_f(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                if (ctx.GetFontProperties(param) == null)
                {
                    //g_warning(_("Font '%i' undefined"), param);
                    return true;
                }

                attr.Font = param;
            }

            return true;
        }

        public static bool doc_fi(RtfDeserialize.ParserContext ctx, RtfState state, int twips)
        {
            if (state is RtfStateAttributes attr)
            {
                string tagName = $"ratifysharp-indent-{twips}";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Indent = Units.ToPixels(RtfDeserialize.TwipsToPango(twips));
                    tag.IndentSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.Indent = twips;
            }

            return true;
        }

        public static bool doc_fs(RtfDeserialize.ParserContext ctx, RtfState state, int halfpoints)
        {
            if (state is RtfStateAttributes attr)
            {
                if (halfpoints <= 0)
                {
                    //g_set_error(error, RTF_ERROR, RTF_ERROR_BAD_FONT_SIZE, _("\\fs%d is invalid, negative or zero font sizes not allowed"), halfpoints);
                    return false;
                }

                double points = halfpoints / 2.0;
                string tagName = string.Format("ratifysharp-fontsize-{0:0.000}", points);
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Size = RtfDeserialize.PointsToPango(points);
                    tag.SizePoints = points;
                    tag.SizeSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.Size = points;
            }

            return true;
        }

        public static bool doc_fsmilli(RtfDeserialize.ParserContext ctx, RtfState state, int milli)
        {
            if (state is RtfStateAttributes attr)
            {
                if (milli <= 0)
                {
                    //g_set_error(error, RTF_ERROR, RTF_ERROR_BAD_FONT_SIZE, _("\\fsmilli%d is invalid, negative or zero font sizes not allowed"), milli);
                    return false;
                }

                double points = milli / 1000.0;
                string tagName = string.Format("ratifysharp-fontsize-{0:0.000}", points);
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Size = RtfDeserialize.PointsToPango(points);
                    tag.SizePoints = points;
                    tag.SizeSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.Size = points;
            }

            return true;
        }

        public static bool doc_highlight(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                if (ctx.ColorTable.Count <= param)
                {
                    //g_set_error(error, RTF_ERROR, RTF_ERROR_UNDEFINED_COLOR, _("Color '%i' undefined"), param);
                    return false;
                }

                string color = ctx.ColorTable[param];
                
                string tagName = $"ratifysharp-highlight-{param}";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.ParagraphBackground = color;
                    tag.SetProperty("paragraph-background-set", new GLib.Value(true));
                    ctx.Tags.Add(tag);
                }

                attr.Background = param;
            }

            return true;
        }
        
        public static bool doc_i(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                string tagName = "ratifysharp-italic";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Style = Pango.Style.Italic;
                    tag.StyleSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.Italic = param != 0;
            }

            return true;
        }

        public static bool doc_lang(RtfDeserialize.ParserContext ctx, RtfState state, int language)
        {
            if (state is RtfStateAttributes attr)
            {
                string tagName = $"ratifysharp-language-{language}";

                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Language = RtfLangCode.LanguageToIsoCode(language);
                    tag.LanguageSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.Language = language;
            }

            return true;
        }

        public static bool doc_li(RtfDeserialize.ParserContext ctx, RtfState state, int twips)
        {
            if (state is RtfStateAttributes attr)
            {
                if (twips < 0)
                    return true; // Silently ignore, not supported in GtkTextBuffer

                string tagName = $"ratifysharp-left-margin-{twips}";

                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.LeftMargin = Units.ToPixels(RtfDeserialize.TwipsToPango(twips));
                    tag.LeftMarginSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.LeftMargin = twips;
            }

            return true;
        }
        
        public static bool doc_ltrch(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                string tagName = "ratifysharp-left-to-right";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Direction = TextDirection.Ltr;
                    ctx.Tags.Add(tag);
                }

                attr.CharDirection = (int)TextDirection.Ltr;
            }

            return true;
        }

        public static bool doc_ltrpar(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                string tagName = "ratifysharp-left-to-right";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Direction = TextDirection.Ltr;
                    ctx.Tags.Add(tag);
                }

                attr.ParDirection = (int)TextDirection.Ltr;
            }

            return true;
        }

        public static bool doc_nosupersub(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                attr.Subscript = false;
                attr.Superscript = false;
            }

            return true;
        }

        public static bool doc_pard(RtfDeserialize.ParserContext ctx, RtfState state, int par)
        {
            if (state is RtfStateAttributes attr)
            {
                attr.Style = -1;
                attr.Justification = -1;
                attr.ParDirection = -1;
                attr.SpaceBefore = 0;
                attr.SpaceAfter = 0;
                attr.IgnoreSpaceBefore = false;
                attr.IgnoreSpaceAfter = false;
                attr.LeftMargin = 0;
                attr.RightMargin = 0;
                attr.Indent = 0;

                attr.Tabs?.Dispose();
                attr.Tabs = null;
            }

            return true;
        }
        
        public static bool doc_plain(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                attr.SetDefaultCharacterAttributes();
                attr.Language = ctx.DefaultLanguage; // override "1024"
            }

            return true;
        }

        public static bool doc_qc(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                string tagName = "ratifysharp-center";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Justification = Justification.Center;
                    tag.JustificationSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.Justification = (int)Justification.Center;
            }

            return true;
        }

        public static bool doc_qj(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                string tagName = "ratifysharp-justified";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Justification = Justification.Fill;
                    tag.JustificationSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.Justification = (int)Justification.Fill;
            }

            return true;
        }

        public static bool doc_ql(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                string tagName = "ratifysharp-left";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Justification = Justification.Left;
                    tag.JustificationSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.Justification = (int)Justification.Left;
            }

            return true;
        }

        public static bool doc_qr(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                string tagName = "ratifysharp-right";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Justification = Justification.Right;
                    tag.JustificationSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.Justification = (int)Justification.Right;
            }

            return true;
        }

        public static bool doc_ri(RtfDeserialize.ParserContext ctx, RtfState state, int twips)
        {
            if (state is RtfStateAttributes attr)
            {
                if (twips < 0)
                    return true; // Silently ignore, not supported in GtkTextBuffer

                string tagName = $"ratifysharp-right-margin-{twips}";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.RightMargin = Units.ToPixels(RtfDeserialize.TwipsToPango(twips));
                    tag.RightMarginSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.RightMargin = twips;
            }

            return true;
        }

        public static bool doc_rtlch(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                string tagName = "ratifysharp-right-to-left";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Direction = TextDirection.Rtl;
                    ctx.Tags.Add(tag);
                }

                attr.CharDirection = (int)TextDirection.Rtl;
            }

            return true;
        }

        public static bool doc_rtlpar(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                string tagName = "ratifysharp-right-to-left";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Direction = TextDirection.Rtl;
                    ctx.Tags.Add(tag);
                }

                attr.ParDirection = (int)TextDirection.Rtl;
            }

            return true;
        }

        public static bool doc_sa(RtfDeserialize.ParserContext ctx, RtfState state, int twips)
        {
            if (state is RtfStateAttributes attr)
            {
                if (twips < 0)
                    return true; // Silently ignore, not supported in GtkTextBuffer

                string tagName = $"ratifysharp-space-after-{twips}";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.PixelsBelowLines = Units.ToPixels(RtfDeserialize.TwipsToPango(twips));
                    tag.PixelsBelowLinesSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.SpaceAfter = twips;
            }

            return true;
        }
        
        public static bool doc_saauto(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                attr.IgnoreSpaceAfter = param != 0;
            }

            return true;
        }

        public static bool doc_sb(RtfDeserialize.ParserContext ctx, RtfState state, int twips)
        {
            if (state is RtfStateAttributes attr)
            {
                if (twips < 0)
                    return true; // Silently ignore, not supported in GtkTextBuffer

                string tagName = $"ratifysharp-space-before-{twips}";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.PixelsAboveLines = Units.ToPixels(RtfDeserialize.TwipsToPango(twips));
                    tag.PixelsAboveLinesSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.SpaceBefore = twips;
            }

            return true;
        }
        
        public static bool doc_sbauto(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                attr.IgnoreSpaceBefore = param != 0;
            }

            return true;
        }

        public static bool doc_scaps(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                string tagName = "ratifysharp-smallcaps";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Variant = Pango.Variant.SmallCaps;
                    tag.VariantSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.SmallCaps = param != 0;
            }

            return true;
        }

        public static bool doc_slleading(RtfDeserialize.ParserContext ctx, RtfState state, int twips)
        {
            if (state is RtfStateAttributes attr)
            {
                if (twips < 0)
                    return true; // Silently ignore, not supported in GtkTextBuffer

                string tagName = $"ratifysharp-leading-{twips}";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.PixelsInsideWrap = Units.ToPixels(RtfDeserialize.TwipsToPango(twips));
                    tag.PixelsInsideWrapSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.Leading = twips;
            }

            return true;
        }

        public static bool doc_strike(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                string tagName = "ratifysharp-strikethrough";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Strikethrough = true;
                    tag.StrikethroughSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.Strikethrough = param != 0;
            }

            return true;
        }

        public static bool doc_sub(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                string tagName = "ratifysharp-subscript";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Rise = RtfDeserialize.PointsToPango(-6);
                    tag.RiseSet = true;
                    tag.Scale = Pango.Scale.XSmall;
                    tag.ScaleSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.Subscript = true;
            }

            return true;
        }

        public static bool doc_super(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                string tagName = "ratifysharp-superscript";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Rise = RtfDeserialize.PointsToPango(6);
                    tag.RiseSet = true;
                    tag.Scale = Pango.Scale.XSmall;
                    tag.ScaleSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.Superscript = true;
            }

            return true;
        }

        public static bool doc_tx(RtfDeserialize.ParserContext ctx, RtfState state, int twips)
        {
            int tabIndex = 0;

            if (state is RtfStateAttributes attr)
            {
                if (attr.Tabs == null)
                {
                    attr.Tabs = new TabArray(1, false);
                }
                else
                {
                    tabIndex = attr.Tabs.Size;
                    attr.Tabs.Resize(tabIndex + 1);
                }

                attr.Tabs.SetTab(tabIndex, TabAlign.Left, RtfDeserialize.TwipsToPango(twips));
            }

            return true;
        }

        public static bool doc_ul(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                string tagName = "ratifysharp-underline-single";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Underline = Underline.Single;
                    tag.UnderlineSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.Underline = param != 0 ? (int)Underline.Single : (int)Underline.None;
            }

            return true;
        }

        public static bool doc_uldb(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                string tagName = "ratifysharp-underline-double";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Underline = Underline.Double;
                    tag.UnderlineSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.Underline = param != 0 ? (int)Underline.Double : (int)Underline.None;
            }

            return true;
        }

        public static bool doc_ulwave(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                string tagName = "ratifysharp-underline-wave";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Underline = Underline.Error;
                    tag.UnderlineSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.Underline = param != 0 ? (int)Underline.Error : (int)Underline.None;
            }

            return true;
        }

        public static bool doc_ulnone(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                attr.Underline = (int)Underline.None;
            }

            return true;
        }

        public static bool doc_ulstyle(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                switch (param & 0xF)
                {
                    case 1:
                        return doc_ul(ctx, attr, 1);
                    case 9:
                        return doc_uldb(ctx, attr, 1);
                }

                return doc_ulnone(ctx, attr, -1);
            }

            return true;
        }

        public static bool doc_up(RtfDeserialize.ParserContext ctx, RtfState state, int halfpoints)
        {
            if (state is RtfStateAttributes attr)
            {
                if (halfpoints != 0)
                {
                    string tagName = $"ratifysharp-up-{halfpoints}";
                    if (ctx.Tags.Lookup(tagName) == null)
                    {
                        var tag = new TextTag(tagName);
                        tag.Rise = RtfDeserialize.HalfPointsToPango(halfpoints);
                        tag.RiseSet = true;
                        ctx.Tags.Add(tag);
                    }
                }

                attr.Rise = halfpoints;
            }

            return true;
        }

        public static bool doc_v(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                string tagName = "ratifysharp-invisible";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    var tag = new TextTag(tagName);
                    tag.Invisible = true;
                    tag.InvisibleSet = true;
                    ctx.Tags.Add(tag);
                }

                attr.Invisible = param != 0;
            }

            return true;
        }

        public static readonly List<RtfDeserialize.ControlWord> FormattedTextControlWords = new List<RtfDeserialize.ControlWord>()
        {
            new RtfDeserialize.ControlWord("b", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_b, 1),
            new RtfDeserialize.ControlWord("cb", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_cb, 0),
            new RtfDeserialize.ControlWord("cf", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_cf, 0),
            new RtfDeserialize.ControlWord("charscalex", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_charscalex, 100),
            new RtfDeserialize.ControlWord("chcbpat", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_cb, 0),
            new RtfDeserialize.ControlWord("dn", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_dn, 6),
            new RtfDeserialize.ControlWord("f", RtfDeserialize.ControlWordType.RequiredParameter, true, doc_f),
            new RtfDeserialize.ControlWord("fi", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_fi, 0),
            new RtfDeserialize.ControlWord("fs", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_fs, 24),
            new RtfDeserialize.ControlWord("fsmilli", RtfDeserialize.ControlWordType.RequiredParameter, true, doc_fsmilli), // Apple extension
            new RtfDeserialize.ControlWord("highlight", RtfDeserialize.ControlWordType.RequiredParameter, true, doc_highlight),
            new RtfDeserialize.ControlWord("i", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_i, 1),
            new RtfDeserialize.ControlWord("lang", RtfDeserialize.ControlWordType.RequiredParameter, true, doc_lang),
            new RtfDeserialize.ControlWord("li", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_li, 0),
            new RtfDeserialize.ControlWord("ltrch", RtfDeserialize.ControlWordType.NoParameter, true, doc_ltrch),
            new RtfDeserialize.ControlWord("ltrpar", RtfDeserialize.ControlWordType.NoParameter, true, doc_ltrpar),
            new RtfDeserialize.ControlWord("nosupersub", RtfDeserialize.ControlWordType.NoParameter, true, doc_nosupersub),
            new RtfDeserialize.ControlWord("pard", RtfDeserialize.ControlWordType.NoParameter, true, doc_pard),
            new RtfDeserialize.ControlWord("plain", RtfDeserialize.ControlWordType.NoParameter, true, doc_plain),
            new RtfDeserialize.ControlWord("qc", RtfDeserialize.ControlWordType.NoParameter, true, doc_qc),
            new RtfDeserialize.ControlWord("qj", RtfDeserialize.ControlWordType.NoParameter, true, doc_qj),
            new RtfDeserialize.ControlWord("ql", RtfDeserialize.ControlWordType.NoParameter, true, doc_ql),
            new RtfDeserialize.ControlWord("qr", RtfDeserialize.ControlWordType.NoParameter, true, doc_qr),
            new RtfDeserialize.ControlWord("ri", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_ri, 0),
            new RtfDeserialize.ControlWord("rtlch", RtfDeserialize.ControlWordType.NoParameter, true, doc_rtlch),
            new RtfDeserialize.ControlWord("rtlpar", RtfDeserialize.ControlWordType.NoParameter, true, doc_rtlpar),
            new RtfDeserialize.ControlWord("sa", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_sa, 0),
            new RtfDeserialize.ControlWord("saauto", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_saauto, 0),
            new RtfDeserialize.ControlWord("sb", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_sb, 0),
            new RtfDeserialize.ControlWord("sbauto", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_sbauto, 0),
            new RtfDeserialize.ControlWord("scaps", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_scaps, 1),
            new RtfDeserialize.ControlWord("slleading", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_slleading, 0), // Apple extension
            new RtfDeserialize.ControlWord("strike", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_strike, 1),
            new RtfDeserialize.ControlWord("sub", RtfDeserialize.ControlWordType.NoParameter, true, doc_sub),
            new RtfDeserialize.ControlWord("super", RtfDeserialize.ControlWordType.NoParameter, true, doc_super),
            new RtfDeserialize.ControlWord("tx", RtfDeserialize.ControlWordType.RequiredParameter, false, doc_tx),
            new RtfDeserialize.ControlWord("ul", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_ul, 1),
            new RtfDeserialize.ControlWord("uld", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_ul, 1), // Treat unsupported types
            new RtfDeserialize.ControlWord("uldash", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_ul, 1), // of underlining as
            new RtfDeserialize.ControlWord("uldashd", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_ul, 1), // regular underlining
            new RtfDeserialize.ControlWord("uldashdd", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_ul, 1),
            new RtfDeserialize.ControlWord("uldb", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_uldb, 1),
            new RtfDeserialize.ControlWord("ulhwave", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_ulwave, 1),
            new RtfDeserialize.ControlWord("ulldash", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_ul, 1),
            new RtfDeserialize.ControlWord("ulnone", RtfDeserialize.ControlWordType.NoParameter, true, doc_ulnone),
            new RtfDeserialize.ControlWord("ulstyle", RtfDeserialize.ControlWordType.RequiredParameter, true, doc_ulstyle), // Apple extension
            new RtfDeserialize.ControlWord("ulth", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_ul, 1),
            new RtfDeserialize.ControlWord("ulthd", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_ul, 1),
            new RtfDeserialize.ControlWord("ulthdash", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_ul, 1),
            new RtfDeserialize.ControlWord("ulthdashd", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_ul, 1),
            new RtfDeserialize.ControlWord("ulthdashdd", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_ul, 1),
            new RtfDeserialize.ControlWord("ulthldash", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_ul, 1),
            new RtfDeserialize.ControlWord("ululdbwave", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_ulwave, 1),
            new RtfDeserialize.ControlWord("ulw", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_ul, 1),
            new RtfDeserialize.ControlWord("ulwave", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_ulwave, 1),
            new RtfDeserialize.ControlWord("up", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_up, 6),
            new RtfDeserialize.ControlWord("v", RtfDeserialize.ControlWordType.OptionalParameter, true, doc_v, 1)
        };

        public static bool doc_chftn(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                ctx.Text += ctx.FootnoteNumber;
            }

            return true;
        }

        public static bool doc_s(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                string tagName = $"ratifysharp-style-{param}";
                if (ctx.Tags.Lookup(tagName) == null)
                {
                    //g_warning(_("Style '%i' undefined"), param);
                    return true;
                }

                attr.Style = param;
            }

            return true;
        }

        public static bool doc_ud(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                attr.UnicodeIgnore = false;
            }

            return true;
        }

        public static bool doc_upr(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfStateAttributes attr)
            {
                attr.UnicodeIgnore = true;
            }

            return true;
        }

        public static RtfDeserialize.DestinationInfo ignore_destination = new RtfIgnore.IgnoreDestinationInfo();

        private static readonly List<RtfDeserialize.ControlWord> _documentTextControlWords = new List<RtfDeserialize.ControlWord>()
        {
            new RtfDeserialize.ControlWord("chftn", RtfDeserialize.ControlWordType.NoParameter, false, doc_chftn),
            new RtfDeserialize.ControlWord("cs", RtfDeserialize.ControlWordType.RequiredParameter, true, doc_s),
            new RtfDeserialize.ControlWord("ds", RtfDeserialize.ControlWordType.RequiredParameter, true, doc_s),
            new RtfDeserialize.ControlWord("nonshppict", RtfDeserialize.ControlWordType.Destination, false, info: ignore_destination),
            new RtfDeserialize.ControlWord("s", RtfDeserialize.ControlWordType.RequiredParameter, true, doc_s),
            new RtfDeserialize.ControlWord("*shppict", RtfDeserialize.ControlWordType.Destination, true, info: ignore_destination),
            new RtfDeserialize.ControlWord("ts", RtfDeserialize.ControlWordType.RequiredParameter, true, doc_s),
            new RtfDeserialize.ControlWord("*ud", RtfDeserialize.ControlWordType.NoParameter, true, doc_ud),
            new RtfDeserialize.ControlWord("upr", RtfDeserialize.ControlWordType.NoParameter, true, doc_upr)
        };

        public static List<RtfDeserialize.ControlWord> DocumentTextControlWords
        {
            get
            {
                var documentTextControlWords = new List<RtfDeserialize.ControlWord>();
                documentTextControlWords.AddRange(SpecialCharacterControlWords);
                documentTextControlWords.AddRange(FormattedTextControlWords);
                documentTextControlWords.AddRange(_documentTextControlWords);
                return documentTextControlWords;
            }
        }

        public static bool doc_rtf(RtfDeserialize.ParserContext ctx, RtfState state, int version)
        {
            if (version != 1)
            {
                //g_set_error(error, RTF_ERROR, RTF_ERROR_BAD_VERSION, _("Unsupported RTF version '%i'"), (gint)version);
                return false;
            }

            return true;
        }

        public static bool doc_ansi(RtfDeserialize.ParserContext ctx, RtfState state, int par)
        {
            ctx.DefaultCodepage = 1252;
            return true;
        }

        public static bool doc_ansicpg(RtfDeserialize.ParserContext ctx, RtfState state, int codepage)
        {
            ctx.Codepage = codepage;
            return true;
        }

        public static bool doc_deff(RtfDeserialize.ParserContext ctx, RtfState state, int font)
        {
            ctx.DefaultFont = font;
            return true;
        }

        public static bool doc_deflang(RtfDeserialize.ParserContext ctx, RtfState state, int language)
        {
            ctx.DefaultLanguage = language;
            return doc_lang(ctx, state, language);
        }

        public static bool doc_footnote(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            // Insert a newline at the end of the document, to separate the coming footnote
            var iter = ctx.TextBuffer.EndIter;
            ctx.TextBuffer.Insert(ref iter, "\n");

            // Move the start and end marks back together
            iter = ctx.TextBuffer.GetIterAtMark(ctx.StartMark);
            ctx.TextBuffer.MoveMark(ctx.EndMark, iter);
            return true;
        }

        public static bool doc_ilvl(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            // Insert n tabs at beginning of line
            var tabString = new string('\t', param);

            var iter = ctx.TextBuffer.EndIter;
            iter.LineOffset = 0;
            ctx.TextBuffer.Insert(ref iter, tabString);

            // Move the start and end marks back together
            ctx.TextBuffer.GetIterAtMark(ctx.StartMark);
            ctx.TextBuffer.MoveMark(ctx.EndMark, iter);
            return true;
        }

        public static bool doc_mac(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            ctx.DefaultCodepage = 10000;
            return true;
        }

        public static bool doc_pc(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            ctx.DefaultCodepage = 437;
            return true;
        }

        public static bool doc_pca(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            ctx.DefaultCodepage = 850;
            return true;
        }
        
        public static RtfDeserialize.DestinationInfo colortbl_destination = new RtfColorTable.ColorTableDestinationInfo();
        public static RtfDeserialize.DestinationInfo field_destination = null; // FIXME: Implement field destination info
        public static RtfDeserialize.DestinationInfo fonttbl_destination = new RtfFontTable.FontTableDestinationInfo();
        public static RtfDeserialize.DestinationInfo footnote_destination = new RtfFootnote.FootnoteDestinationInfo();
        public static RtfDeserialize.DestinationInfo nextgraphic_destination = new RtfPicture.NeXTDestinationInfo();
        public static RtfDeserialize.DestinationInfo pict_destination = new RtfPicture.PictDestinationInfo();
        public static RtfDeserialize.DestinationInfo stylesheet_destination = new RtfStylesheet.StylesheetDestinationInfo();

        private static readonly List<RtfDeserialize.ControlWord> _documentWordTable = new List<RtfDeserialize.ControlWord>()
        {
            new RtfDeserialize.ControlWord("ansi", RtfDeserialize.ControlWordType.NoParameter, false, doc_ansi),
            new RtfDeserialize.ControlWord("ansicpg", RtfDeserialize.ControlWordType.RequiredParameter, false, doc_ansicpg),
            new RtfDeserialize.ControlWord("cell", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\t"), // Fake tables
            new RtfDeserialize.ControlWord("colortbl", RtfDeserialize.ControlWordType.Destination, false, null, 0, null, info: colortbl_destination),
            new RtfDeserialize.ControlWord("deff", RtfDeserialize.ControlWordType.RequiredParameter, false, doc_deff),
            new RtfDeserialize.ControlWord("deflang", RtfDeserialize.ControlWordType.RequiredParameter, false, doc_deflang),
            new RtfDeserialize.ControlWord("field", RtfDeserialize.ControlWordType.Destination, true, null, 0, null, info: field_destination),
            new RtfDeserialize.ControlWord("fonttbl", RtfDeserialize.ControlWordType.Destination, false, null, 0, null, info: fonttbl_destination),
            new RtfDeserialize.ControlWord("footnote", RtfDeserialize.ControlWordType.Destination, true, doc_footnote, 0, null, info: footnote_destination),
            new RtfDeserialize.ControlWord("header", RtfDeserialize.ControlWordType.Destination, false, null, 0, null, info: ignore_destination),
            new RtfDeserialize.ControlWord("ilvl", RtfDeserialize.ControlWordType.RequiredParameter, false, doc_ilvl),
            new RtfDeserialize.ControlWord("info", RtfDeserialize.ControlWordType.Destination, false, null, 0, null, info: ignore_destination),
            new RtfDeserialize.ControlWord("mac", RtfDeserialize.ControlWordType.NoParameter, false, doc_mac),
            new RtfDeserialize.ControlWord("NeXTGraphic", RtfDeserialize.ControlWordType.Destination, false, null, 0, null, info: nextgraphic_destination), // Apple extension
            new RtfDeserialize.ControlWord("pc", RtfDeserialize.ControlWordType.NoParameter, false, doc_pc),
            new RtfDeserialize.ControlWord("pca", RtfDeserialize.ControlWordType.NoParameter, false, doc_pca),
            new RtfDeserialize.ControlWord("pict", RtfDeserialize.ControlWordType.Destination, false, null, 0, null, info: pict_destination),
            new RtfDeserialize.ControlWord("row", RtfDeserialize.ControlWordType.SpecialCharacter, false, null, 0, "\n"), // Fake tables
            new RtfDeserialize.ControlWord("rtf", RtfDeserialize.ControlWordType.RequiredParameter, false, doc_rtf),
            new RtfDeserialize.ControlWord("stylesheet", RtfDeserialize.ControlWordType.Destination, false, null, 0, null, info: stylesheet_destination)
        };

        public static List<RtfDeserialize.ControlWord> DocumentWordTable
        {
            get
            {
                var documentWordTable = new List<RtfDeserialize.ControlWord>();
                documentWordTable.AddRange(DocumentTextControlWords);
                documentWordTable.AddRange(_documentWordTable);
                return documentWordTable;
            }
        }

        public class DocumentDestinationInfo : RtfDeserialize.DestinationInfo
        {
            public override List<RtfDeserialize.ControlWord> ControlWords => DocumentWordTable;

            public override Type StateType => typeof(RtfStateAttributes);

            public override void Cleanup(RtfDeserialize.ParserContext ctx)
            {
            }

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
                if (attr != null && !attr.UnicodeIgnore)
                {
                    TextIter end = ctx.TextBuffer.GetIterAtMark(ctx.EndMark); // shouldn't invalidate end, but it does?
                    ctx.TextBuffer.Insert(ref end, text);

                    TextIter start = ctx.TextBuffer.GetIterAtMark(ctx.StartMark);
                    end = ctx.TextBuffer.GetIterAtMark(ctx.EndMark);

                    ctx.ApplyAttributes(attr, start, end);

                    // Move the two marks back together again
                    ctx.TextBuffer.MoveMark(ctx.StartMark, end);
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

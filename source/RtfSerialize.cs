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
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ratify
{
    public class RtfSerialize
    {
        public class WriterContext
        {
            public TextBuffer TextBuffer;
            public TextIter Start;
            public TextIter End;
            public string Output = string.Empty;
            public TextBuffer LineBuffer;
            public Hashtable TagCodes = new Hashtable(); // Translation table of GtkTextTags to RTF code
            public List<string> FontTable = new List<string>();
            public List<string> ColorTable = new List<string>() { "" }; // Color 0 always black

            /// <summary>
            /// This function is run before processing the actual contents of the buffer. It
            /// generates RTF code for all of the tags in the buffer's tag table, and tells the
            /// context which portion of the text buffer to serialize.
            /// </summary>
            public void AnalyzeBuffer(TextBuffer textBuffer, TextIter start, TextIter end)
            {
                var tagTable = textBuffer.TagTable;
                tagTable.Foreach(ConvertTagToCode);

                TextBuffer = textBuffer;
                Start = start;
                End = end;
            }

            /// <summary>
            /// Return the number of color in the color table. If color is not in the color table, then add it.
            /// </summary>
            public int GetColorFromGdkColor(Color color)
            {
                if (color.Red == 0 && color.Green == 0 && color.Blue == 0)
                    return 0; // Color 0 always black in this implementation

                string colorCode = $"\\red{color.Red >> 8}\\green{color.Green >> 8}\\blue{color.Blue >> 8}";

                int colorNum = 0;

                int link = ColorTable.IndexOf(colorCode);
                if (link == -1)
                {
                    colorNum = ColorTable.Count;
                    ColorTable.Add(colorCode);
                }
                else
                    colorNum = link;

                if (colorNum < 0 || colorNum >= 256)
                    colorNum = 0;

                return colorNum;
            }

            /// <summary>
            /// Generate RTF code for tag, and add it to the context's hashtable of tags to RTF code.
            /// </summary>
            public void ConvertTagToCode(TextTag tag)
            {
                // First check if this is a osxcart named tag that doesn't have a direct
                // Pango attributes equivalent, such as superscript or subscript. Treat these separately.
                string name = (string)tag.GetProperty("name").Val;
                if (name != null)
                {
                    if (name == "ratifysharp-superscript")
                    {
                        TagCodes[tag] = "\\super";
                        return;
                    }
                    else if (name == "ratifysharp-subscript")
                    {
                        TagCodes[tag] = "\\sub";
                        return;
                    }
                }

                // Otherwise, read the attributes one by one and add RTF code for them
                string code = string.Empty;

                bool val = (bool)tag.GetProperty("background-set").Val;
                if (val)
                {
                    var color = (Color)tag.GetProperty("background-gdk").Val;
                    int colorNum = GetColorFromGdkColor(color);
                    code += $"\\chshdng0\\chcbpat{colorNum}\\cb{colorNum}";
                }

                val = (bool)tag.GetProperty("family-set").Val;
                if (val)
                {
                    int fontNum = 0;
                    var family = (string)tag.GetProperty("family").Val;

                    int link = FontTable.IndexOf(family);
                    if (link == -1)
                    {
                        fontNum = FontTable.Count;
                        FontTable.Add(family);
                    }
                    else
                        fontNum = link;

                    code += $"\\f{fontNum}";
                }

                val = (bool)tag.GetProperty("foreground-set").Val;
                if (val)
                {
                    var color = (Color)tag.GetProperty("foreground-gdk").Val;
                    int colorNum = GetColorFromGdkColor(color);
                    code += $"\\cf{colorNum}";
                }

                val = (bool)tag.GetProperty("indent-set").Val;
                if (val)
                {
                    int pixels = (int)tag.GetProperty("indent");
                    code += $"\\fi{PixelsToTwips(pixels)}";
                }

                val = (bool)tag.GetProperty("invisible-set").Val;
                if (val)
                {
                    val = (bool)tag.GetProperty("invisible").Val;
                    if (val)
                        code += "\\v";
                    else
                        code += "\\v0";
                }

                val = (bool)tag.GetProperty("justification-set").Val;
                if (val)
                {
                    var just = (Justification)tag.GetProperty("justification");
                    switch (just)
                    {
                        case Justification.Left:
                            code += "\\ql";
                            break;
                        case Justification.Right:
                            code += "\\qr";
                            break;
                        case Justification.Center:
                            code += "\\qc";
                            break;
                        case Justification.Fill:
                            code += "\\qj";
                            break;
                    }
                }

                val = (bool)tag.GetProperty("language-set").Val;
                if (val)
                {
                    var isoCode = (string)tag.GetProperty("language");
                    code += $"\\lang{RtfLangCode.LanguageToWinCode(isoCode)}";
                }

                val = (bool)tag.GetProperty("left-margin-set").Val;
                if (val)
                {
                    int pixels = (int)tag.GetProperty("left-margin");
                    code += $"\\li{PixelsToTwips(pixels)}";
                }

                val = (bool)tag.GetProperty("paragraph-background-set").Val;
                if (val)
                {
                    var color = (Color)tag.GetProperty("paragraph-background-gdk");
                    int colorNum = GetColorFromGdkColor(color);
                    code += $"\\highlight{colorNum}";
                }

                val = (bool)tag.GetProperty("pixels-above-lines-set").Val;
                if (val)
                {
                    int pixels = (int)tag.GetProperty("pixels-above-lines");
                    code += $"\\sb{PixelsToTwips(pixels)}";
                }

                val = (bool)tag.GetProperty("pixels-below-lines-set").Val;
                if (val)
                {
                    int pixels = (int)tag.GetProperty("pixels-below-lines");
                    code += $"\\sa{PixelsToTwips(pixels)}";
                }

                val = (bool)tag.GetProperty("pixels-inside-wrap-set").Val;
                if (val)
                {
                    int pixels = (int)tag.GetProperty("pixels-inside-wrap");
                    code += $"\\slleading{PixelsToTwips(pixels)}";
                }

                val = (bool)tag.GetProperty("right-margin-set").Val;
                if (val)
                {
                    int pixels = (int)tag.GetProperty("right-margin");
                    code += $"\\ri{PixelsToTwips(pixels)}";
                }

                val = (bool)tag.GetProperty("rise-set").Val;
                if (val)
                {
                    int pango = (int)tag.GetProperty("rise");
                    if (pango > 0)
                        code += $"\\up{PangoToHalfPoints(pango)}";
                    else if (pango < 0)
                        code += $"\\dn{PangoToHalfPoints(-pango)}";
                    else
                        code += "\\up0\\dn0";
                }

                val = (bool)tag.GetProperty("scale-set").Val;
                if (val)
                {
                    double factor = (double)tag.GetProperty("scale");
                    code += $"\\charscalex{(int)(factor * 100)}";
                }

                val = (bool)tag.GetProperty("size-set").Val;
                if (val)
                {
                    double points = (double)tag.GetProperty("size-points");
                    code += $"\\fs{(int)(points * 2)}";

                    // Override with an \fsmilli command if the font size is not a multiple of 1/2 point
                    int milli = (int)(points * 1000);
                    if (milli % 500 != 0)
                        code += $"\\fsmilli{milli}";
                }

                val = (bool)tag.GetProperty("strikethrough-set").Val;
                if (val)
                {
                    val = (bool)tag.GetProperty("strikethrough").Val;
                    if (val)
                        code += "\\strike";
                    else
                        code += "\\strike0";
                }

                val = (bool)tag.GetProperty("style-set").Val;
                if (val)
                {
                    var style = (Pango.Style)tag.GetProperty("style");
                    switch (style)
                    {
                        case Pango.Style.Normal:
                            code += "\\i0";
                            break;
                        case Pango.Style.Oblique:
                        case Pango.Style.Italic:
                            code += "\\i";
                            break;
                    }
                }

                val = (bool)tag.GetProperty("tabs-set").Val;
                if (val)
                {
                    var tabs = (Pango.TabArray)tag.GetProperty("tabs");

                    for (int i = 0; i < tabs.Size; i++)
                    {
                        tabs.GetTab(i, out Pango.TabAlign alignment, out int location);

                        // Alignment can only be LEFT in the current version of Pango
                        int pixels = tabs.PositionsInPixels ? PixelsToTwips(location) : PangoToTwips(location);
                        code += $"\\tx{pixels}";
                    }
                }

                val = (bool)tag.GetProperty("underline-set").Val;
                if (val)
                {
                    var underline = (Pango.Underline)tag.GetProperty("underline");
                    switch (underline)
                    {
                        case Pango.Underline.None:
                            code += "\\ul0\\ulnone";
                            break;
                        case Pango.Underline.Single:
                        case Pango.Underline.Low:
                            code += "\\ul";
                            break;
                        case Pango.Underline.Double:
                            code += "\\uldb";
                            break;
                        case Pango.Underline.Error:
                            code += "\\ulwave";
                            break;
                    }
                }

                val = (bool)tag.GetProperty("variant-set").Val;
                if (val)
                {
                    var variant = (Pango.Variant)tag.GetProperty("variant");
                    switch (variant)
                    {
                        case Pango.Variant.Normal:
                            code += "\\scaps0";
                            break;
                        case Pango.Variant.SmallCaps:
                            code += "\\scaps";
                            break;
                    }
                }

                val = (bool)tag.GetProperty("weight-set").Val;
                if (val)
                {
                    int weight = (int)tag.GetProperty("weight");
                    if (weight >= (int)Pango.Weight.Bold)
                        code += "\\b";
                    else
                        code += "\\b0";
                }

                TagCodes[tag] = code;
            }

            /// <summary>
            /// Write a space to the output buffer if the number of characters output on the
            /// current line is less than 60; otherwise, a newline. If the next space occurs
            /// more than 20 characters further on, the line will still be wider than 80
            /// characters, but this is probably the easiest way to break lines without
            /// looking ahead or backtracking to insert spaces.
            /// </summary>
            public void WriteSpaceOrNewline()
            {
                int currentLineStart = Output.LastIndexOf('\n');
                int lineLength = Output.Substring(currentLineStart).Length;
                Output += lineLength > 60 ? "\n" : " ";
            }


            /// <summary>
            /// This function translates a piece of text, without formatting codes, to RTF.
            /// It replaces special characters by their RTF control word equivalents.
            /// </summary>
            public void WriteRtfText(string text)
            {
                foreach (var ch in text)
                {
                    if (ch == 0x09)
                    {
                        Output += "\\tab";
                    }
                    else if (ch == '\n') // whatever value that is
                    {
                        Output += "\\par";
                    }
                    else if (ch == ' ')
                    {
                        int currentLineStart = Output.LastIndexOf('\n');
                        int lineLength = Output.Substring(currentLineStart).Length;
                        if (lineLength > 60)
                        {
                            Output += "\n";
                        }

                        Output += " ";
                        continue;
                    }
                    else if (ch == 0x5C)
                    {
                        Output += "\\\\";
                        continue;
                    }
                    else if (ch == 0x7B)
                    {
                        Output += "\\{";
                        continue;
                    }
                    else if (ch == 0x7D)
                    {
                        Output += "\\}";
                        continue;
                    }
                    else if (ch > 0x00 && ch < 0x80)
                    {
                        Output += ch;
                        continue;
                    }
                    else if (ch == 0xA0)
                    {
                        Output += "\\~";
                        continue;
                    }
                    else if (ch == 0xAD)
                    {
                        Output += "\\-";
                        continue;
                    }
                    else if (ch >= 0xA1 && ch <= 0xFF)
                    {
                        // Hex character
                        Output += $"\\'{Convert.ToByte(ch).ToString("X2")}";
                        continue;
                    }
                    else if (ch == 0x2002)
                    {
                        Output += "\\enspace";
                    }
                    else if (ch == 0x2003)
                    {
                        Output += "\\emspace";
                    }
                    else if (ch == 0x2005)
                    {
                        Output += "\\qmspace";
                    }
                    else if (ch == 0x200B)
                    {
                        Output += "\\zwbo";
                    }
                    else if (ch == 0x200C)
                    {
                        Output += "\\zwnj";
                    }
                    else if (ch == 0x200D)
                    {
                        Output += "\\zwj";
                    }
                    else if (ch == 0x200E)
                    {
                        Output += "\\ltrmark";
                    }
                    else if (ch == 0x200F)
                    {
                        Output += "\\rtlmark";
                    }
                    else if (ch == 0x2011)
                    {
                        Output += "\\_";
                        continue;
                    }
                    else if (ch == 0x2013)
                    {
                        Output += "\\endash";
                    }
                    else if (ch == 0x2014)
                    {
                        Output += "\\emdash";
                    }
                    else if (ch == 0x2018)
                    {
                        Output += "\\lquote";
                    }
                    else if (ch == 0x2019)
                    {
                        Output += "\\rquote";
                    }
                    else if (ch == 0x201C)
                    {
                        Output += "\\ldblquote";
                    }
                    else if (ch == 0x201D)
                    {
                        Output += "\\rdblquote";
                    }
                    else if (ch == 0x2022)
                    {
                        Output += "\\bullet";
                    }
                    else if (ch == 0x2028)
                    {
                        Output += "\\line";
                    }
                    else
                    {
                        // Unicode character ('?' as fallback)
                        Output += $"\\u{(short)ch}?";
                    }

                    WriteSpaceOrNewline();
                }
            }

            /// <summary>
            /// Analyze a segment of text in which there are no tag flips, but possibly embedded pictures
            /// </summary>
            public void WriteRtfTextAndPictures(TextIter start, TextIter end)
            {
                Pixbuf pixBuf = null;

                TextIter iter = start;
                while (!iter.Equals(end))
                {
                    pixBuf = iter.Pixbuf;
                    if (pixBuf != null)
                        break;

                    iter.ForwardChar();
                }

                if (pixBuf == null)
                {
                    string textEnd = LineBuffer.GetText(start, end, true);
                    WriteRtfText(textEnd);
                    return;
                }

                // Write the text before the pixbuf, insert a \pict destination into the document,
                // and recurse on the text after
                string text = LineBuffer.GetText(start, iter, true);
                WriteRtfText(text);

                var pngBuffer = pixBuf.SaveToBuffer("png", new[] { "compression" }, new[] { "9" });
                if (pngBuffer != null)
                {
                    Output += "{" + $"\\pict\\pngblip\\picw{pixBuf.Width}\\pich{pixBuf.Height}";

                    ulong count = 0;

                    foreach (var b in pngBuffer)
                    {
                        if (count % 40 == 0)
                        {
                            // Newline after every 40 characters
                            Output += "\n";
                        }

                        Output += b.ToString("X2");
                        count++;
                    }

                    Output += "\n}";
                }
                else
                {
                    //g_warning(_("Could not serialize picture, skipping: %s"), error->message);
                }

                iter.ForwardChar();
                WriteRtfTextAndPictures(iter, end);
            }

            /// <summary>
            /// Copy the text paragraph-by-paragraph into a separate buffer
            /// and output each one sequentially with formatting codes
            /// </summary>
            public void WriteRtfParagraphs()
            {
                TextIter lineStart = Start;
                TextIter lineEnd = Start;

                while (lineEnd.InRange(Start, End))
                {
                    // Begin the paragraph by resetting the paragraph properties
                    Output += "{\\pard\\plain";

                    // Get two iterators around the next paragraph of text
                    lineEnd.ForwardToLineEnd();

                    // Skip to the end of any clump of paragraph separators
                    while (lineEnd.EndsLine() && !lineEnd.IsEnd)
                        lineEnd.ForwardChar();

                    if (lineEnd.Compare(End) > 0)
                        lineEnd = End;

                    // Copy the entire paragraph to a separate buffer
                    if (LineBuffer != null)
                        LineBuffer.Dispose();

                    LineBuffer = new TextBuffer(TextBuffer.TagTable);

                    TextIter startIter = LineBuffer.StartIter;
                    TextIter endIter;
                    LineBuffer.InsertRange(ref startIter, lineStart, lineEnd);
                    LineBuffer.GetBounds(out startIter, out endIter);

                    // Insert codes for tags that apply to the whole line, then remove those
                    // tags because we've dealt with them
                    foreach (var tag in lineStart.Tags)
                    {
                        var tagEnd = startIter;
                        tagEnd.ForwardToTagToggle(tag);

                        if (tagEnd.Equals(endIter))
                        {
                            Output += TagCodes[tag];
                            LineBuffer.RemoveTag(tag, startIter, endIter);
                        }
                    }

                    WriteSpaceOrNewline();
                    Output += "{";

                    endIter = startIter;

                    while (!endIter.IsEnd)
                    {
                        // Enclose a section of text without any tag flips between start
                        // and end. Then, make tagStartList a list of tags that open at the
                        // beginning of this section, and tagEndList a list of tags that end
                        // at the end of this section.

                        endIter.ForwardToTagToggle(null);
                        var tagStartList = startIter.GetToggledTags(true).ToList();
                        var tagEndList = endIter.GetToggledTags(false).ToList();
                        var tagOnlyList = new Stack<TextTag>();

                        // Move tags that do not extend before or after this section to tagOnlyList.
                        foreach (var tag in tagStartList)
                        {
                            if (tagEndList.Contains(tag))
                            {
                                tagOnlyList.Push(tag);
                            }
                        }

                        foreach (var tag in tagOnlyList)
                        {
                            tagStartList.Remove(tag);
                            tagEndList.Remove(tag);
                        }

                        int length = Output.Length;

                        // Output the tags in tagStartList
                        foreach (var tag in tagStartList)
                            Output += TagCodes[tag];

                        if (length != Output.Length)
                            WriteSpaceOrNewline();

                        // Output the tags in tagOnlyList, within their own group
                        if (tagOnlyList.Count > 0)
                        {
                            Output += "{";
                            length = Output.Length;

                            foreach (var tag in tagOnlyList)
                                Output += TagCodes[tag];

                            if (length != Output.Length)
                                WriteSpaceOrNewline();
                        }

                        // Output the actual contents of this section
                        WriteRtfTextAndPictures(startIter, endIter);

                        // Close the tagOnlyList group
                        if (tagOnlyList.Count > 0)
                            Output += "}";

                        // If any tags end here, close the group and open another one,
                        // then output the tags that _apply_ to the end iter but do not _start_
                        // there (those will be output in the next iteration and may need to
                        // be in a separate group.)
                        if (tagEndList.Count > 0)
                        {
                            Output += "}{";

                            var tagList = endIter.Tags.ToList();

                            tagStartList = endIter.GetToggledTags(true).ToList();

                            foreach (var tag in tagStartList)
                                tagList.Remove(tag);

                            length = Output.Length;

                            foreach (var tag in tagList)
                                Output += TagCodes[tag];

                            if (length != Output.Length)
                                WriteSpaceOrNewline();
                        }

                        startIter = endIter;
                    }

                    Output += "}}\n";
                    lineStart = lineEnd;
                }
            }

            /// <summary>
            /// Write the RTF header and assorted front matter
            /// </summary>
            public string WriteRtf()
            {
                // Header
                Output += "{\\rtf1\\ansi\\deff0\\uc0\n";

                // Font table
                Output += "{\\fonttbl\n";

                if (FontTable.Count > 0)
                {
                    int i = 0;
                    foreach (var font in FontTable)
                    {
                        var fontNames = font.Split(new[] { ',' }, 2);
                        Output += "{" + $"\\f{i}\\fnil {fontNames[0]};" + "}\n";
                        i++;
                    }
                }
                else
                {
                    // Write at least one font if there are none
                    Output += "{\\f0\\fswiss Sans;}\n";
                }

                Output += "}\n";

                // Color table 
                Output += "{\\colortbl\n";
                foreach (var color in ColorTable)
                {
                    Output += $"{color};\n";
                }
                Output += "}\n";

                // Metadata (provide dummy values because Word will overwrite if missing)
                var assemblyName = GetType().Assembly.GetName();
                Output += "{" + $"\\*\\generator {assemblyName.Name} {assemblyName.Version.ToString(3)}" + "}\n";
                Output += "{\\info {\\author .}{\\company .}{\\title .}\n";

                Output += "{" + $"\\creatim{DateTime.Now.ToString("\\yr%Y\\mo%m\\dy%d\\hr%H\\min%M")}" + "}}\n";
                
                // Preliminary formatting
                Output += $"\\deflang{RtfLangCode.LanguageToWinCode(Pango.Language.Default.ToString())}";
                Output += "\\plain\\widowctrl\\hyphauto\n";

                // Document body
                WriteRtfParagraphs();

                Output += "}";
                return Output;
            }
        }

        // FIXME: these definitions conflate points and pixels

        public static int PixelsToTwips(int pixels)
        {
            return pixels * 20;
        }

        public static int PangoToPoints(int pango)
        {
            return (int)(pango / Pango.Scale.PangoScale);
        }

        public static int PangoToHalfPoints(int pango)
        {
            return (int)(2 * pango / Pango.Scale.PangoScale);
        }

        public static int PangoToTwips(int pango)
        {
            return (int)(20 * pango / Pango.Scale.PangoScale);
        }

        public static string Serialize(TextBuffer contentBuffer, TextIter start, TextIter end)
        {
            var ctx = new WriterContext();

            ctx.AnalyzeBuffer(contentBuffer, start, end);
            return ctx.WriteRtf();
        }

        public static byte[] SerializeTarget(TextBuffer registerBuffer, TextBuffer contentBuffer, TextIter start, TextIter end)
        {
            var ctx = new WriterContext();

            ctx.AnalyzeBuffer(contentBuffer, start, end);

            string rtf = ctx.WriteRtf();
            return Encoding.Unicode.GetBytes(rtf);
        }
    }
}

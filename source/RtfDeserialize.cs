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
using Pango;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace Ratify
{
    public class RtfDeserialize
    {
        public class ParserContext
        {
            public int Codepage = -1;
            public int DefaultCodepage = 1252;
            public int DefaultFont = -1;
            public int DefaultLanguage = 1024;

            public int GroupNestingLevel;
            public Stack<Destination> DestinationStack = new Stack<Destination>();

            public List<string> ColorTable = new List<string>();
            public List<FontProperties> FontTable = new List<FontProperties>();

            public int FootnoteNumber = 1;

            public string RtfText;
            public int Pos;
            public string ConvertBuffer = string.Empty;

            /// <summary>
            /// Text waiting for insertion
            /// </summary>
            public string Text = string.Empty;

            public TextBuffer TextBuffer;
            public TextTagTable Tags;
            public TextMark StartMark;
            public TextMark EndMark;

            public ParserContext(string rtfText, TextBuffer textBuffer, TextIter insert)
            {
                RtfText = rtfText ?? throw new ArgumentNullException("rtfText");
                TextBuffer = textBuffer ?? throw new ArgumentNullException("textBuffer");

                Tags = textBuffer.TagTable;
                StartMark = textBuffer.CreateMark(null, insert, true);
                EndMark = textBuffer.CreateMark(null, insert, false);

                var dest = new Destination();
                dest.Info = new RtfDocument.DocumentDestinationInfo();

                var state = Activator.CreateInstance(dest.Info.StateType) as RtfState;
                dest.StateStack.Push(state);
                DestinationStack.Push(dest);
            }

            /// <summary>
            /// Convenience function to get the current state of the current destination
            /// </summary>
            /// <returns>State attributes object</returns>
            public RtfState GetState()
            {
                var dest = DestinationStack.Peek() as Destination;
                if (dest != null)
                {
                    return dest.StateStack.Peek() as RtfState;
                }

                return null;
            }

            /// <summary>
            /// Returns properties for font numbered index in the font table, or null if such font does not exist
            /// </summary>
            public FontProperties GetFontProperties(int index)
            {
                return FontTable.FirstOrDefault(f => f.Index == index);
            }

            /// <summary>
            /// Parses a control word from the input buffer.
            /// 'word' is the return location for the control word, without a backslash, but with '*' prefixed if the control word is preceded by \*,
            /// which means that the control word represents a destination that should be skipped if it is not recognized.
            /// </summary>
            public bool ParseControlWord(out string word)
            {
                if (RtfText[Pos] != '\\')
                    throw new ArgumentException("RtfText[Pos] == '\\'");

                Pos++;

                if (RtfText[Pos] == '*')
                {
                    // Ignorable destination
                    Pos++;

                    while (char.IsWhiteSpace(RtfText[Pos]))
                        Pos++;

                    if (!ParseControlWord(out string destword))
                    {
                        word = null;
                        return false;
                    }

                    word = "*" + destword;
                }
                else if (char.IsPunctuation(RtfText[Pos]) || RtfText[Pos] == '\n' || RtfText[Pos] == '\r')
                {
                    // Control symbol
                    word = RtfText.Substring(Pos, 1);
                    Pos++;
                }
                else
                {
                    // Control word
                    int length = 0;

                    while (char.IsLetter(RtfText[Pos + length]))
                        length++;

                    if (length == 0)
                    {
                        //g_set_error(error, RTF_ERROR, RTF_ERROR_INVALID_RTF, _("Backslash encountered without control word"));
                        word = null;
                        return false;
                    }

                    word = RtfText.Substring(Pos, length);
                    Pos += length;
                }

                return true;
            }

            /// <summary>
            /// Reads an integer at the current position. If there's no integer at that
            /// position the function returns FALSE, otherwise TRUE. The value is stored in the
            /// location pointed by the 'value' parameter, unless 'value' is NULL. Eat a space
            /// after parsing the number.
            /// </summary>
            public bool ParseIntParameter(out int value)
            {
                // Don't convert the value directly, because it will
                // validate a '+' sign in front of the number, whereas that's not valid
                // according to the RTF spec

                // Find the length of the integer
                int length = 0;

                if (RtfText[Pos] == '-' && char.IsDigit(RtfText[Pos + 1]))
                    length += 2;
                while (char.IsDigit(RtfText[Pos + length]))
                    length++;

                if (length == 0)
                {
                    value = 0;
                    return false;
                }

                // Convert it
                string intstr = RtfText.Substring(Pos, length);
                value = Convert.ToInt32(intstr);

                Pos += length;

                // If the value is delimited by a space, discard the space
                if (RtfText[Pos] == ' ')
                    Pos++;

                return true;
            }


            /// <summary>
            /// Skip one character or control word according to the RTF spec's convoluted skipping rules
            /// </summary>
            public bool SkipCharacterOrControlWord()
            {
                do
                {
                    if (RtfText[Pos] == '{' || RtfText[Pos] == '}')
                        return true; // Skippable data ends before scope delimiter

                    else if (RtfText[Pos] == '\\')
                    {
                        // Special case: \' doesn't follow the regular syntax
                        if (RtfText[Pos + 1] == '\'')
                        {
                            if (!(IsHexDigit(RtfText[Pos + 2]) && IsHexDigit(RtfText[Pos + 3])))
                            {
                                //g_set_error(error, RTF_ERROR, RTF_ERROR_BAD_HEX_CODE, _("Expected a two-character hexadecimal code after \\'"));
                                return false;
                            }

                            Pos += 4;
                            return true;
                        }
                        else
                        {
                            bool success = ParseControlWord(out string word);

                            if (!ParseIntParameter(out int param) && RtfText[Pos] == ' ')
                                Pos++;

                            return success;
                        }
                    }

                    else if (RtfText[Pos] == '\n' || RtfText[Pos] == '\r')
                        Pos++;

                    else
                    {
                        Pos++;
                        return true;
                    }
                } while (true);
            }

            /// <summary>
            /// Carry out the action associated with the control word 'text', as specified
            /// in the current destination's control word table
            /// </summary>
            public bool DoWordAction(string text)
            {
                var dest = DestinationStack.Peek() as Destination;
                var word = dest.Info.ControlWords.FirstOrDefault(w => w.Word == text);

                if (word != null)
                {
                    switch (word.Type)
                    {
                        case ControlWordType.NoParameter:
                            if (RtfText[Pos] == ' ') // Eat a space
                                Pos++;

                            if (word.Action == null)
                                throw new Exception("No action defined");

                            if (word.FlushBuffer)
                                dest.Info.Flush(this);

                            return word.Action(this, GetState());

                        case ControlWordType.OptionalParameter:
                            // If the parameter is optional, carry out the action with the
                            // parameter if there is one, and otherwise with the default parameter
                            if (word.Action == null)
                                throw new Exception("No action defined");

                            if (ParseIntParameter(out int optParam))
                            {
                                if (word.FlushBuffer)
                                    dest.Info.Flush(this);

                                return word.Action(this, GetState(), optParam);
                            }

                            // If no parameter, then eat a space
                            if (RtfText[Pos] == ' ')
                                Pos++;

                            if (word.FlushBuffer)
                                dest.Info.Flush(this);

                            return word.Action(this, GetState(), word.DefaultParam);

                        case ControlWordType.RequiredParameter:
                            if (word.Action == null)
                                throw new Exception("No action defined");

                            if (!ParseIntParameter(out int reqParam))
                            {
                                //g_set_error(error, RTF_ERROR, RTF_ERROR_MISSING_PARAMETER, _("Expected a number after control word '\\%s'"), text);
                                return false;
                            }

                            if (word.FlushBuffer)
                                dest.Info.Flush(this);

                            return word.Action(this, GetState(), reqParam);

                        case ControlWordType.SpecialCharacter:
                            // If the control word represents a special character,
                            // then just insert that character into the buffer
                            if (RtfText[Pos] == ' ') // Eat a space
                                Pos++;

                            if (word.ReplaceText == null)
                                throw new Exception("No replacement text defined");

                            Text += word.ReplaceText;
                            return true;

                        case ControlWordType.Destination:
                            if (RtfText[Pos] == ' ') // Eat a space
                                Pos++;

                            if (word.Action != null && !word.Action(this, GetState()))
                                return false;

                            PushNewDestination(word.DestInfo, null);
                            return true;

                        default:
                            throw new Exception("Invalid word type");
                    }
                }

                // If the control word was not recognized, then ignore it,
                // and any integer parameter that follows
                if (!ParseIntParameter(out int noParam) && RtfText[Pos] == ' ')
                    Pos++;

                // If the control word was an ignorable destination, and was not recognized,
                // push a new "ignore" destination onto the stack
                if (text[0] == '*')
                    PushNewDestination(new RtfIgnore.IgnoreDestinationInfo(), null);

                return true;
            }

            /// <summary>
            /// Push new destination onto the destination stack. If state_to_copy is not
            /// NULL, then initializes the state stack with a copy of that state, otherwise a
            /// blank state.
            /// </summary>
            public void PushNewDestination(DestinationInfo destInfo, RtfState stateToCopy)
            {
                var dest = new Destination();
                dest.Info = destInfo;
                dest.NestingLevel = GroupNestingLevel;

                RtfState stateCopy = null;

                if (stateToCopy != null)
                    stateCopy = stateToCopy.Clone() as RtfState;
                else if (dest.Info.StateType != null)
                    stateCopy = Activator.CreateInstance(dest.Info.StateType) as RtfState;

                dest.StateStack.Push(stateCopy);
                DestinationStack.Push(dest);
            }

            /// <summary>
            /// When exiting a group in the RTF code ('}'), this function is called to pop
            /// one element from the state stack, hence restoring the state before entering the
            /// current group.
            /// </summary>
            public void PopState()
            {
                GroupNestingLevel--;

                var dest = DestinationStack.Peek() as Destination;
                dest.Info.Flush(this);

                if (GroupNestingLevel < dest.NestingLevel)
                {
                    dest.Info.Cleanup(this);

                    dest.StateStack.Pop();
                    DestinationStack.Pop();

                    // Also pop the state of the destination that called this one, since
                    // the opening brace was before the destination control word
                    dest = DestinationStack.Peek() as Destination;
                    dest.Info.Flush(this);
                    dest.StateStack.Pop();
                }
                else
                    dest.StateStack.Pop();
            }

            /// <summary>
            /// When entering a group in the RTF code ('{'), this function copies the current
            /// state and pushes it onto the state stack, so modifications of the state within
            /// the group do not affect the state outside of the group.
            /// </summary>
            public void PushState()
            {
                var dest = DestinationStack.Peek() as Destination;
                dest.Info.Flush(this);

                GroupNestingLevel++;

                var state = dest.StateStack.Peek() as RtfState;
                var stateClone = state.Clone() as RtfState;
                dest.StateStack.Push(stateClone);
            }
            
            /// <summary>
            /// Apply GtkTextTags to the range from start to end, depending on the current attributes 'attr'.
            /// </summary>
            public void ApplyAttributes(RtfStateAttributes attr, TextIter start, TextIter end)
            {
                // Tags with parameters
                if (attr.Style != -1)
                    TextBuffer.ApplyTag($"ratifysharp-style-{attr.Style}", start, end);

                if (attr.Foreground != -1)
                    TextBuffer.ApplyTag($"ratifysharp-foreground-{attr.Foreground}", start, end);

                if (attr.Background != -1)
                    TextBuffer.ApplyTag($"ratifysharp-background-{attr.Background}", start, end);

                if (attr.Highlight != -1)
                    TextBuffer.ApplyTag($"ratifysharp-highlight-{attr.Highlight}", start, end);

                if (attr.Size != 0.0)
                    TextBuffer.ApplyTag(string.Format("ratifysharp-fontsize-{0:0.000}", attr.Size), start, end);

                if (attr.SpaceBefore != 0 && !attr.IgnoreSpaceBefore)
                    TextBuffer.ApplyTag($"ratifysharp-space-before-{attr.SpaceBefore}", start, end);

                if (attr.SpaceAfter != 0 && !attr.IgnoreSpaceAfter)
                    TextBuffer.ApplyTag($"ratifysharp-space-after-{attr.SpaceAfter}", start, end);

                if (attr.LeftMargin != 0)
                    TextBuffer.ApplyTag($"ratifysharp-left-margin-{attr.LeftMargin}", start, end);

                if (attr.RightMargin != 0)
                    TextBuffer.ApplyTag($"ratifysharp-right-margin-{attr.RightMargin}", start, end);

                if (attr.Indent != 0)
                    TextBuffer.ApplyTag($"ratifysharp-indent-{attr.Indent}", start, end);

                if (attr.Invisible)
                    TextBuffer.ApplyTag("ratifysharp-invisible", start, end);

                if (attr.Language != 1024)
                    TextBuffer.ApplyTag($"ratifysharp-language-{attr.Language}", start, end);

                if (attr.Rise != 0)
                    TextBuffer.ApplyTag(string.Format("ratifysharp-{0}-{1}", attr.Rise > 0 ? "up" : "down", (attr.Rise > 0 ? 1 : -1) * attr.Rise), start, end);

                if (attr.Leading != 0)
                    TextBuffer.ApplyTag($"ratifysharp-leading-{attr.Leading}", start, end);

                if (attr.Scale != 100)
                    TextBuffer.ApplyTag($"ratifysharp-scale-{attr.Scale}", start, end);

                // Boolean tags
                if (attr.Italic)
                    TextBuffer.ApplyTag("ratifysharp-italic", start, end);

                if (attr.Bold)
                    TextBuffer.ApplyTag("ratifysharp-bold", start, end);

                if (attr.SmallCaps)
                    TextBuffer.ApplyTag("ratifysharp-smallcaps", start, end);

                if (attr.Strikethrough)
                    TextBuffer.ApplyTag("ratifysharp-strikethrough", start, end);

                if (attr.Underline == (int)Underline.Single)
                    TextBuffer.ApplyTag("ratifysharp-underline-single", start, end);

                if (attr.Underline == (int)Underline.Double)
                    TextBuffer.ApplyTag("ratifysharp-underline-double", start, end);

                if (attr.Underline == (int)Underline.Error)
                    TextBuffer.ApplyTag("ratifysharp-underline-wave", start, end);

                if (attr.Justification == (int)Justification.Left)
                    TextBuffer.ApplyTag("ratifysharp-left", start, end);

                if (attr.Justification == (int)Justification.Right)
                    TextBuffer.ApplyTag("ratifysharp-right", start, end);

                if (attr.Justification == (int)Justification.Center)
                    TextBuffer.ApplyTag("ratifysharp-center", start, end);

                if (attr.Justification == (int)Justification.Fill)
                    TextBuffer.ApplyTag("ratifysharp-justified", start, end);

                if (attr.ParDirection == (int)TextDirection.Rtl)
                    TextBuffer.ApplyTag("ratifysharp-right-to-left", start, end);

                if (attr.ParDirection == (int)TextDirection.Ltr)
                    TextBuffer.ApplyTag("ratifysharp-left-to-right", start, end);

                // Character-formatting direction overrides paragraph formatting
                if (attr.CharDirection == (int)TextDirection.Rtl)
                    TextBuffer.ApplyTag("ratifysharp-right-to-left", start, end);

                if (attr.CharDirection == (int)TextDirection.Ltr)
                    TextBuffer.ApplyTag("ratifysharp-left-to-right", start, end);

                if (attr.Subscript)
                    TextBuffer.ApplyTag("ratifysharp-subscript", start, end);

                if (attr.Superscript)
                    TextBuffer.ApplyTag("ratifysharp-superscript", start, end);

                // Special
                if (attr.Font != -1)
                    TextBuffer.ApplyTag($"ratifysharp-font-{attr.Font}", start, end);

                else if (DefaultFont != -1 && FontTable.Count > DefaultFont)
                    TextBuffer.ApplyTag($"ratifysharp-font-{DefaultFont}", start, end);

                if (attr.Tabs != null)
                {
                    // Create a separate tag for each Pango.TabArray
                    string tagName = $"ratifysharp-tabs-{attr.Tabs.Handle.ToInt64()}";

                    TextTag tag = Tags.Lookup(tagName);
                    if (tag == null)
                    {
                        tag = new TextTag(tagName);
                        tag.Tabs = attr.Tabs;
                        tag.TabsSet = true;
                        Tags.Add(tag);
                    }

                    TextBuffer.ApplyTag(tag, start, end);
                }
            }

            /// <summary>
            /// Insert picture into text buffer at current insertion mark.
            /// </summary>
            /// <param name="pixbuf">Picture</param>
            public void InsertPictureIntoTextBuffer(Pixbuf pixbuf)
            {
                TextIter iter = TextBuffer.GetIterAtMark(EndMark);
                TextBuffer.InsertPixbuf(ref iter, pixbuf);
            }

            /// <summary>
            /// Convert the character ch to unicode and add to the context's buffer
            /// </summary>
            public bool ConvertHexCharToUnicode(char ch)
            {
                // Determine the character encoding that ch is in. First see if the current
                // destination diverts us to another codepage (e.g., \fcharset in the \fonttbl
                // destination) and if not, use either the current codepage or the default codepage.
                var dest = DestinationStack.Peek() as Destination;

                int codepage = dest.Info.GetCodepage(this);
                if (codepage == -1)
                    codepage = Codepage;

                var enc = GetEncodingForCodepage(codepage);
                if (enc == null)
                    enc = GetEncodingForCodepage(DefaultCodepage);

                if (enc == null)
                {
                    //g_set_error(error, RTF_ERROR, RTF_ERROR_UNSUPPORTED_CHARSET, _("Character set %d is not supported"), (DefaultCodepage == -1) ? codepage : DefaultCodepage);
                    return false;
                }

                // Now see if there was any incompletely converted text left over from previous characters
                string textToConvert = string.Empty;
                if (ConvertBuffer.Length > 0)
                {
                    ConvertBuffer += ch;
                    textToConvert = ConvertBuffer;
                    ConvertBuffer = string.Empty;
                }
                else
                    textToConvert = new string(ch, 1);

                try
                {
                    // Should use question mark as a replacement (fallback) character
                    var convertedBytes = Encoding.Convert(enc, Encoding.UTF8, enc.GetBytes(textToConvert));
                    var convertedText = Encoding.UTF8.GetString(convertedBytes);

                    Text += convertedText;
                }
                catch (Exception)
                {
                    // If there is a "partial input" error, then save the text
                    // in the convert buffer and retrieve it if there is another
                    // consecutive \'xx code
                    //if (converterror->code == G_CONVERT_ERROR_PARTIAL_INPUT)
                    //    g_string_append(ctx->convertbuffer, text_to_convert);
                    //else
                    //    g_warning(_("Conversion error: %s"), converterror->message);
                    //g_clear_error(&converterror);
                }

                return true;
            }

            /// <summary>
            /// The main parser loop
            /// </summary>
            /// <returns>Success or failure</returns>
            public bool ParseRtf()
            {
                do
                {
                    if (RtfText[Pos] == '\0')
                    {
                        //g_set_error(error, RTF_ERROR, RTF_ERROR_MISSING_BRACE, _("File ended unexpectedly"));
                        return false;
                    }

                    if (RtfText[Pos] == '{')
                    {
                        Pos++;
                        PushState();
                    }
                    else if (RtfText[Pos] == '}')
                    {
                        Pos++;
                        PopState();
                    }
                    else if (RtfText[Pos] == '\\')
                    {
                        // Special case: \' doesn't follow the regular syntax
                        if (RtfText[Pos + 1] == '\'')
                        {
                            if (!(IsHexDigit(RtfText[Pos + 2]) && IsHexDigit(RtfText[Pos + 3])))
                            {
                                //g_set_error(error, RTF_ERROR, RTF_ERROR_BAD_HEX_CODE, _("Expected a two-character hexadecimal code after \\'"));
                                return false;
                            }

                            string hexcode = RtfText.Substring(Pos + 2, 2);
                            char ch = ConvertHexToChar(hexcode);

                            Pos += 4;

                            if (!ConvertHexCharToUnicode(ch))
                                return false;
                        }
                        else
                        {
                            bool success = ParseControlWord(out string word) && DoWordAction(word);
                            if (!success)
                                return false;
                        }
                    }
                    else if (RtfText[Pos] == '\n' || RtfText[Pos] == '\r')
                    {
                        // Ignore newlines
                        Pos++;
                    }
                    else if (RtfText[Pos] > 0xFF)
                    {
                        // Ignore high characters (they should be encoded with \'xx)
                        Pos++;
                    }
                    else
                    {
                        // If there is any partial wide character in the convert buffer, then
                        // try to combine it with this one as a double-byte character
                        if (ConvertBuffer.Length > 0)
                        {
                            if (!ConvertHexCharToUnicode(RtfText[Pos]))
                                return false;
                        }
                        else
                            // Add character to current string
                            Text += RtfText[Pos];

                        Pos++;
                    }

                } while (GroupNestingLevel > 0);

                return true;
            }
        }
        
        public static bool Deserialize(TextBuffer contentBuffer, TextIter iter, string data)
        {
            if (!data.StartsWith("{\\rtf"))
            {
                //g_set_error(error, RTF_ERROR, RTF_ERROR_INVALID_RTF, _("RTF format must begin with '{\\rtf'"));
                return false;
            }

            var ctx = new ParserContext(data, contentBuffer, iter);
            return ctx.ParseRtf();
        }

        public static bool DeserializeTarget(TextBuffer registerBuffer, TextBuffer contentBuffer, TextIter iter, byte[] data, ulong length, bool createTags)
        {
            var dataStr = Encoding.ASCII.GetString(data);
            if (!dataStr.StartsWith("{\\rtf"))
            {
                //g_set_error(error, RTF_ERROR, RTF_ERROR_INVALID_RTF, _("RTF format must begin with '{\\rtf'"));
                return false;
            }

            var ctx = new ParserContext(dataStr, contentBuffer, iter);
            return ctx.ParseRtf();
        }

        /// <summary>
        /// Get the encoding for a codepage or return null.
        /// </summary>
        /// <param name="codepage">Codepage</param>
        /// <returns>Encoding or null</returns>
        public static Encoding GetEncodingForCodepage(int codepage)
        {
            if (codepage == -1)
                return null;

            try
            {
                var enc = Encoding.GetEncoding(codepage);
                return enc;
            }
            catch { }

            return null;
        }

        public static bool IsHexDigit(char c)
        {
            if ('0' <= c && c <= '9') return true;
            if ('a' <= c && c <= 'f') return true;
            if ('A' <= c && c <= 'F') return true;
            return false;
        }

        public static char ConvertHexToChar(string hex)
        {
            try
            {
                return (char)int.Parse(hex, NumberStyles.HexNumber);
            }
            catch { }

            return char.MinValue;
        }

        /// <summary>
        /// Converts the string to an array of bytes.
        /// </summary>
        /// <param name="hex">Hex-String</param>
        /// <returns>byte-Array</returns>
        public static byte[] ConvertHexStringToByteArray(string hex)
        {
            return Enumerable.Range(0, hex.Length)
                             .Where(x => x % 2 == 0)
                             .Select(x => Convert.ToByte(hex.Substring(x, 2), 16))
                             .ToArray();
        }

        public static int PointsToPango(double pts)
        {
            return (int)(pts * Pango.Scale.PangoScale);
        }

        public static int HalfPointsToPango(int halfpts)
        {
            return (int)(halfpts * Pango.Scale.PangoScale / 2);
        }

        public static int TwipsToPango(int twips)
        {
            return (int)(twips * Pango.Scale.PangoScale / 20);
        }

        public abstract class DestinationInfo
        {
            public abstract List<ControlWord> ControlWords { get; }
            public abstract Type StateType { get; }
            public abstract void Cleanup(ParserContext ctx);
            public abstract void Flush(ParserContext ctx);
            public abstract int GetCodepage(ParserContext ctx);
        }

        public enum ControlWordType
        {
            NoParameter,
            OptionalParameter,
            RequiredParameter,
            SpecialCharacter,
            Destination
        }

        public delegate bool ControlWordAction(ParserContext ctx, RtfState state, int par = 0);

        public class ControlWord
        {
            public string Word;
            public ControlWordType Type;
            public bool FlushBuffer;
            public ControlWordAction Action;
            public int DefaultParam;
            public string ReplaceText;
            public DestinationInfo DestInfo;

            public ControlWord(string word, ControlWordType type, bool flushBuffer, ControlWordAction action = null, int defaultParam = 0, string replaceText = "", DestinationInfo info = null)
            {
                Word = word;
                Type = type;
                FlushBuffer = flushBuffer;
                Action = action;
                DefaultParam = defaultParam;
                ReplaceText = replaceText;
                DestInfo = info;
            }
        }

        public class Destination
        {
            public int NestingLevel;
            public Stack<RtfState> StateStack = new Stack<RtfState>();
            public DestinationInfo Info;
        }

        public class FontProperties
        {
            public int Index;
            public int Codepage;
            public string FontName;
        }
    }
}

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
    public static class RtfColorTable
    {
        public class RtfColorTableState : RtfState
        {
            public int Red = 0;
            public int Green = 0;
            public int Blue = 0;

            public override object Clone()
            {
                return MemberwiseClone();
            }

            public override void Dispose()
            {
            }
        }
        

        public static bool ct_red(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfColorTableState colorState)
                colorState.Red = param;

            return true;
        }

        public static bool ct_green(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfColorTableState colorState)
                colorState.Green = param;

            return true;
        }

        public static bool ct_blue(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfColorTableState colorState)
                colorState.Blue = param;

            return true;
        }

        public static readonly List<RtfDeserialize.ControlWord> ColorTableWordTable = new List<RtfDeserialize.ControlWord>()
        {
            new RtfDeserialize.ControlWord("red", RtfDeserialize.ControlWordType.RequiredParameter, true, ct_red),
            new RtfDeserialize.ControlWord("green", RtfDeserialize.ControlWordType.RequiredParameter, true, ct_green),
            new RtfDeserialize.ControlWord("blue", RtfDeserialize.ControlWordType.RequiredParameter, true, ct_blue)
        };

        public class ColorTableDestinationInfo : RtfDeserialize.DestinationInfo
        {
            public override List<RtfDeserialize.ControlWord> ControlWords => ColorTableWordTable;

            public override Type StateType => typeof(RtfColorTableState);

            public override void Cleanup(RtfDeserialize.ParserContext ctx)
            {
            }

            /// <summary>
            /// If the text contains a semicolon,
            /// add the RGB code to the color table and reset the color table state.
            /// </summary>
            public override void Flush(RtfDeserialize.ParserContext ctx)
            {
                if (ctx == null)
                    throw new ArgumentNullException("ctx");

                var state = ctx.GetState() as RtfColorTableState;
                if (state != null)
                {
                    if (ctx.Text.Contains(";"))
                    {
                        var color = $"{state.Red:X2}{state.Green:X2}{state.Blue:X2}";
                        ctx.ColorTable.Add(color);

                        state.Red = 0;
                        state.Green = 0;
                        state.Blue = 0;
                    }

                    ctx.Text = string.Empty;
                }
            }
            
            public override int GetCodepage(RtfDeserialize.ParserContext ctx)
            {
                return 0;
            }
        }
    }
}

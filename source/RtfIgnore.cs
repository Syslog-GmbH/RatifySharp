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
    public static class RtfIgnore
    {
        public class RtfIgnoreState : RtfState
        {
            public override object Clone()
            {
                return MemberwiseClone();
            }

            public override void Dispose()
            {
            }
        }

        public static readonly List<RtfDeserialize.ControlWord> IgnoreWordTable = new List<RtfDeserialize.ControlWord>()
        {
        };
        
        public class IgnoreDestinationInfo : RtfDeserialize.DestinationInfo
        {
            public override List<RtfDeserialize.ControlWord> ControlWords => IgnoreWordTable;

            public override Type StateType => typeof(RtfIgnoreState);

            public override void Cleanup(RtfDeserialize.ParserContext ctx)
            {
            }

            public override void Flush(RtfDeserialize.ParserContext ctx)
            {
                if (ctx == null)
                    throw new ArgumentNullException("ctx");

                ctx.Text = string.Empty;
            }

            public override int GetCodepage(RtfDeserialize.ParserContext ctx)
            {
                return 0;
            }
        }
    }
}

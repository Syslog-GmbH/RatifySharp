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

using Gdk;
using Pango;

namespace Ratify
{
    /// <summary>
    /// All destinations dealing with inserting graphics into the document.
    /// </summary>
    public static class RtfPicture
    {
        public enum PictType
        {
            PICT_TYPE_EMF,
            PICT_TYPE_PNG,
            PICT_TYPE_JPEG,
            PICT_TYPE_MAC,
            PICT_TYPE_OS2,
            PICT_TYPE_WMF,
            PICT_TYPE_DIB,
            PICT_TYPE_BMP
        }

        public class RtfPictState : RtfState
        {
            public PictType Type = PictType.PICT_TYPE_WMF;
            public int TypeParam = 1;
            public PixbufLoader Loader;
            public bool Error = false;

            public int Width = -1;
            public int Height = -1;
            public int WidthGoal = -1;
            public int HeightGoal = -1;
            public int XScale = 100;
            public int YScale = 100;

            public override object Clone()
            {
                // FIXME: Clone PixbufLoader??
                return MemberwiseClone();
            }

            public override void Dispose()
            {
                // FIXME: Dispose PixbufLoader??
            }

            /// <summary>
            /// Send a message to the GdkPixbufLoader to change its expected size
            /// after parsing a width or height declaration.
            /// </summary>
            public void AdjustLoaderSize()
            {
                if (Loader != null && (Width != -1 || WidthGoal != -1) && (Height != -1 || HeightGoal != -1))
                {
                    Loader.SetSize(
                        (WidthGoal != -1) ? WidthGoal : Width,
                        (HeightGoal != -1) ? HeightGoal : Height);
                }
            }
        }

        public class RtfNeXTState : RtfState
        {
            public int Width = -1;
            public int Height = -1;

            public override object Clone()
            {
                return MemberwiseClone();
            }

            public override void Dispose()
            {
            }
        }

        public static bool pic_dibitmap(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfPictState pict)
            {
                if (param != 0)
                {
                    //g_set_error(error, RTF_ERROR, RTF_ERROR_BAD_PICT_TYPE, _("Invalid bitmap type '%i' for \\dibitmap"), param);
                    return false;
                }

                pict.Type = PictType.PICT_TYPE_DIB;
                pict.TypeParam = 0;
            }

            return true;
        }

        public static bool pic_emfblip(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfPictState pict)
            {
                pict.Type = PictType.PICT_TYPE_EMF;
            }

            return true;
        }

        public static bool pic_jpegblip(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfPictState pict)
            {
                pict.Type = PictType.PICT_TYPE_JPEG;
            }

            return true;
        }

        public static bool pic_macpict(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfPictState pict)
            {
                pict.Type = PictType.PICT_TYPE_MAC;
            }

            return true;
        }

        public static bool pic_pich(RtfDeserialize.ParserContext ctx, RtfState state, int pixels)
        {
            if (state is RtfPictState pict)
            {
                pict.Height = pixels;
                pict.AdjustLoaderSize();
            }

            return true;
        }

        public static bool pic_pichgoal(RtfDeserialize.ParserContext ctx, RtfState state, int twips)
        {
            if (state is RtfPictState pict)
            {
                pict.HeightGoal = Units.ToPixels(RtfDeserialize.TwipsToPango(twips));
                pict.AdjustLoaderSize();
            }

            return true;
        }

        public static bool pic_picscalex(RtfDeserialize.ParserContext ctx, RtfState state, int percent)
        {
            if (state is RtfPictState pict)
            {
                pict.XScale = percent;
            }

            return true;
        }

        public static bool pic_picscaley(RtfDeserialize.ParserContext ctx, RtfState state, int percent)
        {
            if (state is RtfPictState pict)
            {
                pict.YScale = percent;
            }

            return true;
        }

        public static bool pic_picw(RtfDeserialize.ParserContext ctx, RtfState state, int pixels)
        {
            if (state is RtfPictState pict)
            {
                pict.Width = pixels;
                pict.AdjustLoaderSize();
            }

            return true;
        }

        public static bool pic_picwgoal(RtfDeserialize.ParserContext ctx, RtfState state, int twips)
        {
            if (state is RtfPictState pict)
            {
                pict.WidthGoal = Units.ToPixels(RtfDeserialize.TwipsToPango(twips));
                pict.AdjustLoaderSize();
            }

            return true;
        }

        public static bool pic_pmmetafile(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfPictState pict)
            {
                pict.Type = PictType.PICT_TYPE_OS2;
                pict.TypeParam = param;
            }

            return true;
        }

        public static bool pic_pngblip(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfPictState pict)
            {
                pict.Type = PictType.PICT_TYPE_PNG;
            }

            return true;
        }

        public static bool pic_wbitmap(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfPictState pict)
            {
                if (param != 0)
                {
                    //g_set_error(error, RTF_ERROR, RTF_ERROR_BAD_PICT_TYPE, _("Invalid bitmap type '%i' for \\wbitmap"), param);
                    return false;
                }

                pict.Type = PictType.PICT_TYPE_BMP;
                pict.TypeParam = 0;
            }

            return true;
        }

        public static bool pic_wmetafile(RtfDeserialize.ParserContext ctx, RtfState state, int param)
        {
            if (state is RtfPictState pict)
            {
                pict.Type = PictType.PICT_TYPE_WMF;
                pict.TypeParam = param;
            }

            return true;
        }
        
        public static readonly List<RtfDeserialize.ControlWord> PictWordTable = new List<RtfDeserialize.ControlWord>()
        {
            new RtfDeserialize.ControlWord("dibitmap", RtfDeserialize.ControlWordType.RequiredParameter, false, pic_dibitmap),
            new RtfDeserialize.ControlWord("emfblip", RtfDeserialize.ControlWordType.NoParameter, false, pic_emfblip),
            new RtfDeserialize.ControlWord("jpegblip", RtfDeserialize.ControlWordType.NoParameter, false, pic_jpegblip),
            new RtfDeserialize.ControlWord("macpict", RtfDeserialize.ControlWordType.NoParameter, false, pic_macpict),
            new RtfDeserialize.ControlWord("pich", RtfDeserialize.ControlWordType.RequiredParameter, false, pic_pich),
            new RtfDeserialize.ControlWord("pichgoal", RtfDeserialize.ControlWordType.RequiredParameter, false, pic_pichgoal),
            new RtfDeserialize.ControlWord("picscalex", RtfDeserialize.ControlWordType.OptionalParameter, false, pic_picscalex, 100),
            new RtfDeserialize.ControlWord("picscaley", RtfDeserialize.ControlWordType.OptionalParameter, false, pic_picscaley, 100),
            new RtfDeserialize.ControlWord("picw", RtfDeserialize.ControlWordType.RequiredParameter, false, pic_picw),
            new RtfDeserialize.ControlWord("picwgoal", RtfDeserialize.ControlWordType.RequiredParameter, false, pic_picwgoal),
            new RtfDeserialize.ControlWord("pmmetafile", RtfDeserialize.ControlWordType.RequiredParameter, false, pic_pmmetafile),
            new RtfDeserialize.ControlWord("pngblip", RtfDeserialize.ControlWordType.NoParameter, false, pic_pngblip),
            new RtfDeserialize.ControlWord("wbitmap", RtfDeserialize.ControlWordType.RequiredParameter, false, pic_wbitmap),
            new RtfDeserialize.ControlWord("wmetafile", RtfDeserialize.ControlWordType.OptionalParameter, false, pic_wmetafile, 1)
        };

        public class PictDestinationInfo : RtfDeserialize.DestinationInfo
        {
            public override List<RtfDeserialize.ControlWord> ControlWords => PictWordTable;

            public override Type StateType => typeof(RtfPictState);

            public override void Cleanup(RtfDeserialize.ParserContext ctx)
            {
                if (ctx == null)
                    throw new ArgumentNullException("ctx");

                var state = ctx.GetState() as RtfPictState;
                if (state == null || state.Error || state.Loader == null)
                    return;

                if (!state.Loader.Close())
                {
                    //g_warning(_("Error closing pixbuf loader: %s"), error->message);
                }

                var picture = state.Loader.Pixbuf;
                if (picture != null)
                {
                    // Scale picture if needed
                    if (state.XScale != 100 || state.YScale != 100)
                    {
                        int newwidth = picture.Width * state.XScale / 100;
                        int newheight = picture.Height * state.YScale / 100;
                        Pixbuf newpicture = picture.ScaleSimple(newwidth, newheight, InterpType.Bilinear);
                        picture.Dispose();
                        picture = newpicture;
                    }

                    ctx.InsertPictureIntoTextBuffer(picture);
                }
                else
                {
                    //g_warning(_("Error loading picture"));
                }
            }

            public override void Flush(RtfDeserialize.ParserContext ctx)
            {
                if (ctx == null)
                    throw new ArgumentNullException("ctx");

                var mimeTypes = new[] {
                    "image/x-emf", "image/png", "image/jpeg", "image/x-pict",
                    "OS/2 Presentation Manager", "image/x-wmf", "image/x-bmp", "image-x-bmp"
                }; // "OS/2 Presentation Manager" isn't supported

                var state = ctx.GetState() as RtfPictState;
                if (state == null || state.Error)
                    return;

                if (ctx.Text.Length == 0)
                    return;

                // If no GdkPixbufLoader has been initialized yet, then do that
                if (state.Loader == null)
                {
                    // Make sure the MIME type we want to load is present in the list of
                    // formats compiled into our GdkPixbuf library
                    foreach (var format in Pixbuf.Formats)
                    {
                        if (state.Loader != null)
                            break;

                        foreach (var mime in format.MimeTypes)
                        {
                            if (string.Compare(mime, mimeTypes[(int)state.Type], true) == 0)
                            {
                                state.Loader = new PixbufLoader(mimeTypes[(int)state.Type]);
                                if (state.Loader == null)
                                {
                                    //g_warning(_("Error loading picture of MIME type '%s': %s"), mimetypes[state->type], error->message);
                                    state.Error = true;
                                }
                                break;
                            }
                        }
                    }
                    
                    if (state.Loader == null && !state.Error)
                    {
                        //g_warning(_("Module for loading MIME type '%s' not found"), mimetypes[state->type]);
                        state.Error = true;
                    }

                    if (state.Error)
                        return;

                    state.AdjustLoaderSize();
                }

                // Convert the "text" into binary data
                var writeBuffer = RtfDeserialize.ConvertHexStringToByteArray(ctx.Text);

                // Write the "text" into the GdkPixbufLoader
                if (!state.Loader.Write(writeBuffer))
                {
                    //g_warning(_("Error reading \\pict data: %s"), error->message);
                    state.Error = true;
                }
                
                ctx.Text = string.Empty;
            }

            public override int GetCodepage(RtfDeserialize.ParserContext ctx)
            {
                return 0;
            }
        }

        public static readonly List<RtfDeserialize.ControlWord> ShppictWordTable = new List<RtfDeserialize.ControlWord>()
        {
            new RtfDeserialize.ControlWord("pict", RtfDeserialize.ControlWordType.Destination, false, info: new PictDestinationInfo())
        };

        public class ShppictDestinationInfo : RtfDeserialize.DestinationInfo
        {
            public override List<RtfDeserialize.ControlWord> ControlWords => ShppictWordTable;

            public override Type StateType => null;

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
        
        public static bool ng_height(RtfDeserialize.ParserContext ctx, RtfState state, int twips)
        {
            if (state is RtfNeXTState next)
            {
                next.Height = Units.ToPixels(RtfDeserialize.TwipsToPango(twips));
            }

            return true;
        }

        public static bool ng_width(RtfDeserialize.ParserContext ctx, RtfState state, int twips)
        {
            if (state is RtfNeXTState next)
            {
                next.Width = Units.ToPixels(RtfDeserialize.TwipsToPango(twips));
            }

            return true;
        }

        public static readonly List<RtfDeserialize.ControlWord> NeXTWordTable = new List<RtfDeserialize.ControlWord>()
        {
            new RtfDeserialize.ControlWord("height", RtfDeserialize.ControlWordType.RequiredParameter, false, ng_height),
            new RtfDeserialize.ControlWord("width", RtfDeserialize.ControlWordType.RequiredParameter, false, ng_width)
        };

        public class NeXTDestinationInfo : RtfDeserialize.DestinationInfo
        {
            public override List<RtfDeserialize.ControlWord> ControlWords => NeXTWordTable;

            public override Type StateType => typeof(RtfNeXTState);

            public override void Cleanup(RtfDeserialize.ParserContext ctx)
            {
                var state = ctx.GetState() as RtfNeXTState;
                if (state == null)
                    return;

                // Load the file from the filename in the pending text buffer
                string fileName = ctx.Text.Trim();
                ctx.Text = string.Empty;
                
                var pixbuf = new Pixbuf(fileName, state.Width, state.Height, false);
                if (pixbuf == null)
                {
                    //g_warning(_("Error loading picture from file '%s': %s"), fileName, error->message);
                    return;
                }

                ctx.InsertPictureIntoTextBuffer(pixbuf);
            }

            public override void Flush(RtfDeserialize.ParserContext ctx)
            {
                // Ignore text, but leave it in the pending text buffer, because we use it in Cleanup()
            }

            public override int GetCodepage(RtfDeserialize.ParserContext ctx)
            {
                return 65001; // UTF-8
            }
        }
    }
}

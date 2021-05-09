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
using System.Linq;

namespace Ratify
{
    public static class RtfLangCode
    {
        public class IsoLangCode
        {
            public int WinCode;
            public string IsoCode;

            public IsoLangCode(int winCode, string isoCode)
            {
                WinCode = winCode;
                IsoCode = isoCode;
            }
        }

        // These are from the RTF spec.
        // The ISO codes are the best I could figure out from the Wikipedia page.
        private static readonly IsoLangCode[] _isolangcodes =
        {
            new IsoLangCode(0x0400, "zxx"), /* None */
            new IsoLangCode(0x0401, "ar-sa"), /* Arabic (Saudi Arabia) */
            new IsoLangCode(0x0402, "bg"), /* Bulgarian */
            new IsoLangCode(0x0403, "ca"), /* Catalan */
            new IsoLangCode(0x0404, "zh-tw"), /* Chinese (Taiwan) */
            new IsoLangCode(0x0405, "cs"), /* Czech */
            new IsoLangCode(0x0406, "da"), /* Danish */
            new IsoLangCode(0x0407, "de-de"), /* German (Germany) */
            new IsoLangCode(0x0408, "el"), /* Greek */
            new IsoLangCode(0x0409, "en-us"), /* English (United States) */
            new IsoLangCode(0x040A, "es-es"), /* Spanish (Spain, Traditional Sort)? */
            new IsoLangCode(0x040B, "fi"), /* Finnish */
            new IsoLangCode(0x040C, "fr-fr"), /* French (France) */
            new IsoLangCode(0x040D, "he"), /* Hebrew */
            new IsoLangCode(0x040E, "hu"), /* Hungarian */
            new IsoLangCode(0x040F, "is"), /* Icelandic */
            new IsoLangCode(0x0410, "it-it"), /* Italian (Italy) */
            new IsoLangCode(0x0411, "ja"), /* Japanese */
            new IsoLangCode(0x0412, "ko"), /* Korean */
            new IsoLangCode(0x0413, "nl-nl"), /* Dutch (Netherlands) */
            new IsoLangCode(0x0414, "nb"), /* Norwegian, Bokmal */
            new IsoLangCode(0x0415, "pl"), /* Polish */
            new IsoLangCode(0x0416, "pt-br"), /* Portuguese (Brazil) */
            new IsoLangCode(0x0417, "rm"), /* Romansh */
            new IsoLangCode(0x0418, "ro-ro"), /* Romanian (Romania) */
            new IsoLangCode(0x0419, "ru-ru"), /* Russian (Russia) */
            new IsoLangCode(0x041A, "hr-hr"), /* Croatian (Croatia) */
            new IsoLangCode(0x041B, "sk"), /* Slovak */
            new IsoLangCode(0x041C, "sq"), /* Albanian */
            new IsoLangCode(0x041D, "sv-se"), /* Swedish (Sweden) */
            new IsoLangCode(0x041E, "th"), /* Thai */
            new IsoLangCode(0x041F, "tr"), /* Turkish */
            new IsoLangCode(0x0420, "ur-pk"), /* Urdu (Islamic Republic of Pakistan) */
            new IsoLangCode(0x0421, "id"), /* Indonesian */
            new IsoLangCode(0x0422, "uk"), /* Ukrainian */
            new IsoLangCode(0x0423, "be"), /* Belarussian */
            new IsoLangCode(0x0424, "sl"), /* Slovenian */
            new IsoLangCode(0x0425, "et"), /* Estonian */
            new IsoLangCode(0x0426, "lv"), /* Latvian */
            new IsoLangCode(0x0427, "lt"), /* Lithuanian */
            new IsoLangCode(0x0428, "tg"), /* Tajik */
            new IsoLangCode(0x0429, "fa"), /* Persian */
            new IsoLangCode(0x042A, "vi"), /* Vietnamese */
            new IsoLangCode(0x042B, "hy"), /* Armenian */
            new IsoLangCode(0x042C, "az"), /* Azerbaijani (Latin) */
            new IsoLangCode(0x042D, "eu"), /* Basque */
            new IsoLangCode(0x042E, "hsb"), /* Upper Sorbian */
            new IsoLangCode(0x042F, "mk"), /* Macedonian */
            //new IsoLangCode(0x0430, Sutu (South Africa) -- has no code? */
            new IsoLangCode(0x0431, "ts"), /* Tsonga */
            new IsoLangCode(0x0432, "tn"), /* Setswana */
            new IsoLangCode(0x0433, "ve"), /* Venda */
            new IsoLangCode(0x0434, "xh"), /* isiXhosa */
            new IsoLangCode(0x0435, "zu"), /* isiZulu */
            new IsoLangCode(0x0436, "af"), /* Afrikaans */
            new IsoLangCode(0x0437, "ka"), /* Georgian */
            new IsoLangCode(0x0438, "fo"), /* Faroese */
            new IsoLangCode(0x0439, "hi"), /* Hindi */
            new IsoLangCode(0x043A, "mt"), /* Maltese */
            new IsoLangCode(0x043B, "se-no"), /* Northern Sami (Norway) */
            new IsoLangCode(0x043C, "gd"), /* Gaelic (Scotland) */
            new IsoLangCode(0x043D, "yi"), /* Yiddish */
            new IsoLangCode(0x043E, "ms-my"), /* Malay (Malaysia) */
            new IsoLangCode(0x043F, "kk"), /* Kazakh */
            new IsoLangCode(0x0440, "ky"), /* Kyrgyz */
            new IsoLangCode(0x0441, "sw"), /* Kiswahili */
            new IsoLangCode(0x0442, "tk"), /* Turkmen */
            new IsoLangCode(0x0443, "uz"), /* Uzbek (Latin) */
            new IsoLangCode(0x0444, "tt"), /* Tatar */
            new IsoLangCode(0x0445, "bn-in"), /* Bengali (India) */
            new IsoLangCode(0x0446, "pa-in"), /* Punjabi (India) */
            new IsoLangCode(0x0447, "gu"), /* Gujarati */
            new IsoLangCode(0x0448, "or"), /* Oriya */
            new IsoLangCode(0x0449, "ta"), /* Tamil */
            new IsoLangCode(0x044A, "te"), /* Telugu */
            new IsoLangCode(0x044B, "kn"), /* Kannada */
            new IsoLangCode(0x044C, "ml"), /* Malayalam */
            new IsoLangCode(0x044D, "as"), /* Assamese */
            new IsoLangCode(0x044E, "mr"), /* Marathi */
            new IsoLangCode(0x044F, "sa"), /* Sanskrit */
            new IsoLangCode(0x0450, "mn-mn"), /* Mongolian (Cyrillic, Mongolia) */
            new IsoLangCode(0x0451, "bo"), /* Tibetan */
            new IsoLangCode(0x0452, "cy"), /* Welsh */
            new IsoLangCode(0x0453, "km"), /* Khmer */
            new IsoLangCode(0x0454, "lo"), /* Lao */
            new IsoLangCode(0x0455, "my"), /* Burmese */
            new IsoLangCode(0x0456, "gl"), /* Galician */
            new IsoLangCode(0x0457, "kok"), /* Konkani */
            new IsoLangCode(0x0458, "mni"), /* Manipuri */
            new IsoLangCode(0x0459, "sd"), /* Sindhi (Devanagari) */
            new IsoLangCode(0x045A, "syr"), /* Syriac */
            new IsoLangCode(0x045B, "si"), /* Sinhala */
            new IsoLangCode(0x045C, "chr"), /* Cherokee */
            new IsoLangCode(0x045D, "iu"), /* Inuktitut (Syllabics) */
            new IsoLangCode(0x045E, "am"), /* Amharic */
            new IsoLangCode(0x045F, "tia-ma"), /* Tamazight (Arabic, Morocco) */
            new IsoLangCode(0x0460, "ks"), /* Kashmiri (Arabic) */
            new IsoLangCode(0x0461, "ne-np"), /* Nepali (Nepal) */
            new IsoLangCode(0x0462, "fy"), /* Frisian */
            new IsoLangCode(0x0463, "ps"), /* Pashto */
            new IsoLangCode(0x0464, "fil"), /* Filipino */
            new IsoLangCode(0x0465, "dv"), /* Divehi */
            new IsoLangCode(0x0466, "bin"), /* Edo */
            new IsoLangCode(0x0467, "ff"), /* Fulfulde */
            new IsoLangCode(0x0468, "ha"), /* Hausa */
            new IsoLangCode(0x0469, "ibb"), /* Ibibio */
            new IsoLangCode(0x046A, "yo"), /* Yoruba */
            new IsoLangCode(0x046B, "qu-bo"), /* Quechua (Bolivia) */
            new IsoLangCode(0x046C, "nso"), /* Sotho sa Leboa */
            new IsoLangCode(0x046D, "ba"), /* Bashkir */
            new IsoLangCode(0x046E, "lb"), /* Luxembourgish */
            new IsoLangCode(0x046F, "kl"), /* Greenlandic */
            new IsoLangCode(0x0470, "ig"), /* Igbo */
            new IsoLangCode(0x0471, "kr"), /* Kanuri */
            new IsoLangCode(0x0472, "om"), /* Oromo */
            new IsoLangCode(0x0473, "ti-et"), /* Tigrinya (Ethiopia) */
            new IsoLangCode(0x0474, "gn"), /* Guarani */
            new IsoLangCode(0x0475, "haw"), /* Hawaiian */
            new IsoLangCode(0x0476, "la"), /* Latin */
            new IsoLangCode(0x0477, "so"), /* Somali */
            new IsoLangCode(0x0478, "iii"), /* Yi */
            new IsoLangCode(0x0479, "pap"), /* Papiamentu */
            new IsoLangCode(0x047A, "arn"), /* Mapudungun */
            new IsoLangCode(0x047C, "moh"), /* Mohawk */
            new IsoLangCode(0x047E, "br"), /* Breton */
            new IsoLangCode(0x0480, "ug"), /* Uighur */
            new IsoLangCode(0x0481, "mi"), /* Maori */
            new IsoLangCode(0x0482, "oc"), /* Occitan */
            new IsoLangCode(0x0483, "co"), /* Corsican */
            new IsoLangCode(0x0484, "gsw"), /* Alsatian? */
            new IsoLangCode(0x0485, "sah"), /* Yakut */
            new IsoLangCode(0x0486, "quc"), /* K'iche' */
            new IsoLangCode(0x0487, "rw"), /* Kinyarwanda */
            new IsoLangCode(0x0488, "wo"), /* Wolof */
            new IsoLangCode(0x048C, "prs"), /* Dari? */
            new IsoLangCode(0x0801, "ar-iq"), /* Arabic (Iraq) */
            new IsoLangCode(0x0804, "zh-cn"), /* Chinese (P.R.C.) */
            new IsoLangCode(0x0807, "de-ch"), /* German (Switzerland) */
            new IsoLangCode(0x0809, "en-gb"), /* English (United Kingdom) */
            new IsoLangCode(0x080A, "es-mx"), /* Spanish (Mexico) */
            new IsoLangCode(0x080C, "fr-be"), /* French (Belgium) */
            new IsoLangCode(0x0810, "it-ch"), /* Italian (Switzerland) */
            new IsoLangCode(0x0813, "nl-be"), /* Dutch (Belgium) */
            new IsoLangCode(0x0814, "nn"), /* Norwegian, Nynorsk */
            new IsoLangCode(0x0816, "pt-pt"), /* Portuguese (Portugal) */
            new IsoLangCode(0x0818, "ro-md"), /* Romanian (Moldova) */
            new IsoLangCode(0x0819, "ru-md"), /* Russian (Moldova) */
            new IsoLangCode(0x081A, "sr-rs"), /* Serbian (Latin, Serbia) */
            new IsoLangCode(0x081D, "sv-fi"), /* Swedish (Finland) */
            new IsoLangCode(0x0820, "ur-in"), /* Urdu (India) */
            new IsoLangCode(0x0827, "lt"), /* Lithuanian Traditional? */
            new IsoLangCode(0x082C, "az"), /* Azerbaijani (Cyrillic) */
            new IsoLangCode(0x082E, "dsb"), /* Lower Sorbian */
            new IsoLangCode(0x083B, "se-se"), /* Northern Sami (Sweden) */
            new IsoLangCode(0x083C, "ga"), /* Gaelic (Ireland) */
            new IsoLangCode(0x083E, "ms-bn"), /* Malay (Brunei Darussalam) */
            new IsoLangCode(0x0843, "uz"), /* Uzbek (Cyrillic) */
            new IsoLangCode(0x0845, "bn-bd"), /* Bengali (Bangladesh) */
            new IsoLangCode(0x0846, "pa-pk"), /* Punjabi (Pakistan) */
            new IsoLangCode(0x0850, "mn-cn"), /* Mongolian (Traditional Mongolian, P.R.C.) */
            new IsoLangCode(0x0851, "dz"), /* Dzongkha */
            new IsoLangCode(0x0859, "sd"), /* Sindhi (Arabic) */
            new IsoLangCode(0x085D, "iu"), /* Inuktitut (Latin) */
            new IsoLangCode(0x085F, "tia-dz"), /* Tamazight (Latin, Algeria) */
            new IsoLangCode(0x0860, "ks"), /* Kashmiri */
            new IsoLangCode(0x0861, "ne-in"), /* Nepali (India) */
            new IsoLangCode(0x086B, "qu-ec"), /* Quechua (Ecuador) */
            new IsoLangCode(0x0873, "ti-er"), /* Tigrinya (Eritrea) */
            new IsoLangCode(0x0C00, "und"), /* "Custom Current"? */
            new IsoLangCode(0x0C01, "ar-eg"), /* Arabic (Egypt) */
            new IsoLangCode(0x0C04, "zh-hk"), /* Chinese (Hong Kong S.A.R.) */
            new IsoLangCode(0x0C07, "de-at"), /* German (Austria) */
            new IsoLangCode(0x0C09, "en-au"), /* English (Australia) */
            new IsoLangCode(0x0C0A, "es-es"), /* Spanish (Spain, International Sort)? */
            new IsoLangCode(0x0C0C, "fr-ca"), /* French (Canada) */
            new IsoLangCode(0x0C1A, "sr-rs"), /* Serbian (Cyrillic, Serbia) */
            new IsoLangCode(0x0C3B, "se-fi"), /* Northern Sami (Finland) */
            new IsoLangCode(0x0C6B, "qu-pe"), /* Quechua (Peru) */
            new IsoLangCode(0x1001, "ar-ly"), /* Arabic (Libya) */
            new IsoLangCode(0x1004, "zh-sg"), /* Chinese (Singapore) */
            new IsoLangCode(0x1007, "de-lu"), /* German (Luxembourg) */
            new IsoLangCode(0x1009, "en-ca"), /* English (Canada) */
            new IsoLangCode(0x100A, "es-gt"), /* Spanish (Guatemala) */
            new IsoLangCode(0x100C, "fr-ch"), /* French (Switzerland) */
            new IsoLangCode(0x101A, "hr-ba"), /* Croatian (Bosnia and Herzegovina) */
            new IsoLangCode(0x103B, "smj-no"), /* Lule Sami (Norway) */
            new IsoLangCode(0x1401, "ar-dz"), /* Arabic (Algeria) */
            new IsoLangCode(0x1404, "zh-mo"), /* Chinese (Macao S.A.R.) */
            new IsoLangCode(0x1407, "de-li"), /* German (Liechtenstein) */
            new IsoLangCode(0x1409, "en-nz"), /* English (New Zealand) */
            new IsoLangCode(0x140A, "es-cr"), /* Spanish (Costa Rica) */
            new IsoLangCode(0x140C, "fr-lu"), /* French (Luxembourg) */
            new IsoLangCode(0x141A, "bs"), /* Bosnian (Latin) */
            new IsoLangCode(0x143B, "smj-se"), /* Lule Sami (Sweden) */
            new IsoLangCode(0x1801, "ar-ma"), /* Arabic (Morocco) */
            new IsoLangCode(0x1809, "en-ie"), /* English (Ireland) */
            new IsoLangCode(0x180A, "es-pa"), /* Spanish (Panama) */
            new IsoLangCode(0x180C, "fr-mc"), /* French (Monaco) */
            new IsoLangCode(0x181A, "sr-ba"), /* Serbian (Latin, Bosnia and Herzegovina) */
            new IsoLangCode(0x183B, "sma-no"), /* Southern Sami (Norway) */
            new IsoLangCode(0x1C01, "ar-tn"), /* Arabic (Tunisia) */
            new IsoLangCode(0x1C09, "en-za"), /* English (South Africa) */
            new IsoLangCode(0x1C0A, "es-do"), /* Spanish (Dominican Republic) */
            new IsoLangCode(0x1C0C, "fr"), /* French (West Indies)? */
            new IsoLangCode(0x1C1A, "sr-ba"), /* Serbian (Cyrillic, Bosnia and Herzegovina) */
            new IsoLangCode(0x1C3B, "sma-se"), /* Southern Sami (Sweden) */
            new IsoLangCode(0x2001, "ar-om"), /* Arabic (Oman) */
            new IsoLangCode(0x2009, "en-jm"), /* English (Jamaica) */
            new IsoLangCode(0x200A, "es-ve"), /* Spanish (Venezuela) */
            new IsoLangCode(0x200C, "fr-re"), /* French (Reunion) */
            new IsoLangCode(0x201A, "bs"), /* Bosnian (Cyrillic) */
            new IsoLangCode(0x203B, "sms"), /* Skolt Sami */
            new IsoLangCode(0x2401, "ar-ye"), /* Arabic (Yemen) */
            new IsoLangCode(0x2409, "en"), /* English (Caribbean)? */
            new IsoLangCode(0x240A, "es-co"), /* Spanish (Colombia) */
            new IsoLangCode(0x240C, "fr-cd"), /* French (Democratic Republic of Congo) */
            new IsoLangCode(0x243B, "smn"), /* Inari Sami */
            new IsoLangCode(0x2801, "ar-sy"), /* Arabic (Syria) */
            new IsoLangCode(0x2809, "en-bz"), /* English (Belize) */
            new IsoLangCode(0x280A, "es-pe"), /* Spanish (Peru) */
            new IsoLangCode(0x280C, "fr-sn"), /* French (Senegal) */
            new IsoLangCode(0x2C01, "ar-jo"), /* Arabic (Jordan) */
            new IsoLangCode(0x2C09, "en-tt"), /* English (Trinidad and Tobago) */
            new IsoLangCode(0x2C0A, "es-ar"), /* Spanish (Argentina) */
            new IsoLangCode(0x2C0C, "fr-cm"), /* French (Cameroon) */
            new IsoLangCode(0x3001, "ar-lb"), /* Arabic (Lebanon) */
            new IsoLangCode(0x3009, "en-zw"), /* English (Zimbabwe) */
            new IsoLangCode(0x300A, "es-ec"), /* Spanish (Ecuador) */
            new IsoLangCode(0x300C, "fr-ci"), /* French (Cote d'Ivoire) */
            new IsoLangCode(0x3401, "ar-kw"), /* Arabic (Kuwait) */
            new IsoLangCode(0x3409, "en-ph"), /* English (Philippines) */
            new IsoLangCode(0x340A, "es-cl"), /* Spanish (Chile) */
            new IsoLangCode(0x340C, "fr-ml"), /* French (Mali) */
            new IsoLangCode(0x3801, "ar-ae"), /* Arabic (U.A.E.) */
            new IsoLangCode(0x3809, "en-id"), /* English (Indonesia) */
            new IsoLangCode(0x380A, "es-uy"), /* Spanish (Uruguay) */
            new IsoLangCode(0x380C, "fr-ma"), /* French (Morocco) */
            new IsoLangCode(0x3C01, "ar-bh"), /* Arabic (Bahrain) */
            new IsoLangCode(0x3C09, "en-hk"), /* English (Hong Kong S.A.R.) */
            new IsoLangCode(0x3C0A, "es-py"), /* Spanish (Paraguay) */
            new IsoLangCode(0x3C0C, "fr-ht"), /* French (Haiti) */
            new IsoLangCode(0x4001, "ar-qa"), /* Arabic (Qatar) */
            new IsoLangCode(0x4009, "en-in"), /* English (India) */
            new IsoLangCode(0x400A, "es-bo"), /* Spanish (Bolivia) */
            new IsoLangCode(0x4409, "en-my"), /* English (Malaysia) */
            new IsoLangCode(0x440A, "es-sv"), /* Spanish (El Salvador) */
            new IsoLangCode(0x4809, "en-sg"), /* English (Singapore) */
            new IsoLangCode(0x480A, "es-hn"), /* Spanish (Honduras) */
            new IsoLangCode(0x4C0A, "es-ni"), /* Spanish (Nicaragua) */
            new IsoLangCode(0x500A, "es-pr"), /* Spanish (Puerto Rico) */
            new IsoLangCode(0x540A, "es-us"), /* Spanish (United States) */
            // Aliases
            new IsoLangCode(0x0000, "zxx"), /* None */
            new IsoLangCode(0x0009, "en"), /* English? */
            new IsoLangCode(0x0013, "nl"), /* "Dutch Preferred"? */
            new IsoLangCode(0x0409, "c") /* C/POSIX locale = English? */
        };

        // Do not free return value
        public static string LanguageToIsoCode(int wincode)
        {
            var lang = _isolangcodes.FirstOrDefault(i => i.WinCode == wincode);
            if (lang != null)
            {
                return lang.IsoCode;
            }

            return "zxx"; // No language
        }

        public static int LanguageToWinCode(string isocode)
        {
            var lang = _isolangcodes.FirstOrDefault(i => i.IsoCode.Equals(isocode, StringComparison.OrdinalIgnoreCase));
            if (lang != null)
            {
                return lang.WinCode;
            }

            return 1024; // No language
        }
    }
}

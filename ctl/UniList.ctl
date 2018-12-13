VERSION 5.00
Begin VB.UserControl UniList 
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ForeColor       =   &H80000008&
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "UniList.ctx":0000
End
Attribute VB_Name = "UniList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************************************
'* UniList 0.9.1 - Unicode listbox user control
'* --------------------------------------------
'* By Vesa Piittinen aka Merri, http://vesa.piittinen.name/ <vesa@piittinen.name>
'* Unicode on 2000/XP/Vista
'*
'* LICENSE
'* -------
'* http://creativecommons.org/licenses/by-sa/1.0/fi/deed.en
'*
'* Terms: 1) If you make your own version, share using this same license.
'*        2) When used in a program, mention my name in the program's credits.
'*        3) May not be used as a part of commercial (unicode) controls suite.
'*        4) Free for any other commercial and non-commercial usage.
'*        5) Use at your own risk. No support guaranteed.
'*
'* SUPPORT FOR UNICONTROLS
'* -----------------------
'* http://www.vbforums.com/showthread.php?t=500026
'* http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=69738&lngWId=1
'*
'* REQUIREMENTS
'* ------------
'* Note: TLBs are compiled to your program so you don't need to distribute the files
'* - OleGuids3.tlb      = Ole Guid and interface definitions 3.0
'* - UniListModule.bas
'*
'* NOTES
'* -----
'* This control is not stop button safe. This is because of IOleInPlaceActiveObject.
'* Currently there is no IDE safe solution for it, yet it is very important for this project.
'* You can press the pause button, but don't push stop!
'*
'* HOW TO ADD TO YOUR PROGRAM
'* --------------------------
'* 1) OPTIONAL: Copy OleGuids3.tlb to Windows system folder.
'* 2) Copy UniListModule.bas, UniList.ctl and UniList.ctx to your project folder.
'* 3) In your project, add a reference to OleGuids3.tlb (Project > References...)
'* 4) Add UniListModule.bas
'* 5) Add UniList.ctl
'*
'* VERSION HISTORY
'* ---------------
'* Version 0.9.1 RELEASE CANDIDATE 1 (2008-06-14)
'* - Many new properties and methods plus some similar fixes/changes as in the UniText control.
'*
'* Version 0.6 BETA (2008-06-11)
'* - Initial release.
'*
'* CREDITS
'* -------
'* - Mike Gainer, Matt Curland and Bill Storage for their work on IOLEInPlaceActivate
'* - Paul Caton and LaVolpe for their work on SelfSub, SelfHook and SelfCallback
'*************************************************************************************************
Option Explicit

Public Event Click(Button As UniListMouseButton)
Attribute Click.VB_MemberFlags = "200"
Public Event DblClick(Button As UniListMouseButton)
Public Event FontChanged()
Public Event ItemCheck(Index As Long)
Public Event KeyDown(KeyCode As Integer, ByVal Shift As UniListShift)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, ByVal Shift As UniListShift)
Public Event MouseDown(Button As UniListMouseButton, ByVal Shift As UniListShift, X As Single, Y As Single)
Public Event MouseEnter()
Public Event MouseLeave()
Public Event MouseMove(Button As UniListMouseButton, ByVal Shift As UniListShift, X As Single, Y As Single)
Public Event MouseUp(Button As UniListMouseButton, ByVal Shift As UniListShift, X As Single, Y As Single)
Public Event MouseWheel(ByVal Wheel As UniListMouseWheel, ByVal Shift As UniListShift)
'Public Event OLECompleteDrag(Effect As Long)
'Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
'Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
'Public Event OLESetData(Data As DataObject, DataFormat As Integer)
'Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Public Event Scroll(ByVal Direction As UniListScrollDirection)

Private Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Private Type MSGFILTER
    NMHDR As NMHDR
    Msg As Long
    wParam As Long
    lParam As Long
End Type

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Public Enum UniListAppearance
    [Classic 3D]
    [Windows 3D]
End Enum

Public Enum UniListBorderStyle
    [No Border] = 0&
    [Flat3D] = 1&
    [3D] = 2&
End Enum

Public Enum UniListLocale
    [Invalid Locale] = 0&
    [Locale Afrikaans] = &H436  ' Afrikaans (South Africa)    af-ZA   Latn    1252
    [Locale Albanian] = &H41C  ' Albanian (Albania)  sq-AL   Latn    1252
    [Locale Alsatian] = &H484  ' Windows Vista and later: Alsatian (France)  gsw-FR
    [Locale Amharic] = &H45E  ' Windows Vista and later: Amharic (Ethiopia)     am-ET       Unicode only
    [Locale Arabic/Algeria] = &H1401 ' Arabic (Algeria)    ar-DZ   Arab    1256
    [Locale Arabic/Bahrain] = &H3C01 ' Arabic (Bahrain)    ar-BH   Arab    1256
    [Locale Arabic/Egypt] = &HC01  ' Arabic (Egypt)  ar-EG   Arab    1256
    [Locale Arabic/Iraq] = &H801  ' Arabic (Iraq)   ar-IQ   Arab    1256
    [Locale Arabic/Jordan] = &H2C01 ' Arabic (Jordan)     ar-JO   Arab    1256
    [Locale Arabic/Kuwait] = &H3401 ' Arabic (Kuwait)     ar-KW   Arab    1256
    [Locale Arabic/Lebanon] = &H3001 '' Arabic (Lebanon)    ar-LB   Arab    1256
    [Locale Arabic/Libya] = &H1001 '' Arabic (Libya)  ar-LY   Arab    1256
    [Locale Arabic/Morocco] = &H1801 '' Arabic (Morocco)    ar-MA   Arab    1256
    [Locale Arabic/Oman] = &H2001 '' Arabic (Oman)   ar-OM   Arab    1256
    [Locale Arabic/Qatar] = &H4001 '' Arabic (Qatar)  ar-QA   Arab    1256
    [Locale Arabic/Saudia Arabia] = &H401  '' Arabic (Saudi Arabia)   ar-SA   Arab    1256
    [Locale Arabic/Syria] = &H2801 '' Arabic (Syria)  ar-SY   Arab    1256
    [Locale Arabic/Tunisia] = &H1C01 '' Arabic (Tunisia)    ar-TN   Arab    1256
    [Locale Arabic/U.A.E.] = &H3801 '' Arabic (U.A.E.)     ar-AE   Arab    1256
    [Locale Arabic/Yemen] = &H2401 '' Arabic (Yemen)  ar-YE   Arab    1256
    [Locale Armenian] = &H42B  '' Windows 2000 and later: Armenian (Armenia)  hy-AM   Armn    Unicode only
    [Locale Assamese] = &H44D  '' Windows Vista and later: Assamese (India)   as-IN       Unicode only
    [Locale Azeri/Cyrillic] = &H82C  '' Azeri (Azerbaijan, Cyrillic)    az-Cyrl-AZ  Cyrl    1251
    [Locale Azeri/Latin] = &H42C  '' Azeri (Azerbaijan, Latin)   az-Latn-AZ  Latn    1254
    [Locale Bashkir] = &H46D  '' Windows Vista and later: Bashkir (Russia)   ba-RU
    [Locale Basque] = &H42D  '' Basque (Basque)     eu-ES   Latn    1252
    [Locale Belarusian] = &H423  '' Belarusian (Belarus)    be-BY   Cyrl    1251
    [Locale Bengali] = &H445  '' Windows XP SP2 and later: Bengali (India)   bn-IN   Beng    Unicode only
    [Locale Bosnian/Cyrillic] = &H201A '' Windows XP SP2 and later (downloadable); Windows Vista and later: Bosnian (Bosnia and Herzegovina, Cyrillic)    bs-Cyrl-BA  Cyrl    1251
    [Locale Bosnian/Latin] = &H141A '' Windows XP SP2 and later: Bosnian (Bosnia and Herzegovina, Latin)   bs-Latn-BA  Latn    1250
    [Locale Breton] = &H47E  '' Breton (France)     br-FR   Latn    1252
    [Locale Bulgarian] = &H402  '' Bulgarian (Bulgaria)    bg-BG   Cyrl    1251
    [Locale Burmese] = &H455  '' Not supported: Burmese
    [Locale Catalan] = &H403  '' Catalan (Catalan)   ca-ES   Latn    1252
    [Locale Chinese/Hong Kong SAR, PRC] = &HC04  '' Chinese (Hong Kong SAR, PRC)    zh-HK   Hant    950
    [Locale Chinese/Macao SAR] = &H1404 '' Windows 98/Me, Windows XP and later: Chinese (Macao SAR)    zh-MO   Hant    950
    [Locale Chinese/PRC] = &H804  '' Chinese (PRC)   zh-CN   Hans    936
    [Locale Chinese/Singapore] = &H1004 '' Chinese (Singapore)     zh-SG   Hans    936
    [Locale Chinese/Taiwan] = &H404  '' Chinese (Taiwan)    zh-TW   Hant    950
    [Locale Croatian/Bosnia and Herzegovina/Latin] = &H101A '' Windows XP SP2 and later: Croatian (Bosnia and Herzegovina, Latin)  hr-BA   Latn    1250
    [Locale Croatian] = &H41A  '' Croatian (Croatia)  hr-HR   Latn    1250
    [Locale Czech] = &H405  '' Czech (Czech Republic)  cs-CZ   Latn    1250
    [Locale Danish] = &H406  '' Danish (Denmark)    da-DK   Latn    1252
    [Locale Dari] = &H48C  '' Windows XP and later: Dari (Afghanistan)    prs-AF  Arab    1256
    [Locale Divehi] = &H465  '' Windows XP and later: Divehi (Maldives)     dv-MV   Thaa    Unicode only
    [Locale Dutch/Belgium] = &H813  '' Dutch (Belgium)     nl-BE   Latn    1252
    [Locale Dutch/Netherlands] = &H413  '' Dutch (Netherlands)     nl-NL   Latn    1252
    [Locale English/Australia] = &HC09  '' English (Australia)     en-AU   Latn    1252
    [Locale English/Belize] = &H2809 '' English (Belize)    en-BZ   Latn    1252
    [Locale English/Canada] = &H1009 '' English (Canada)    en-CA   Latn    1252
    [Locale English/Caribbean] = &H2409 '' English (Caribbean)     en-029  Latn    1252
    [Locale English/India] = &H4009 '' Windows Vista and later: English (India)    en-IN   Latn    1252
    [Locale English/Ireland] = &H1809 '' English (Ireland)   en-IE   Latn    1252
    [Locale English/Jamaica] = &H2009 '' English (Jamaica)   en-JM   Latn    1252
    [Locale English/Malaysia] = &H4409 '' Windows Vista and later: English (Malaysia)     en-MY   Latn    1252
    [Locale English/New Zealand] = &H1409 '' English (New Zealand)   en-NZ   Latn    1252
    [Locale English/Philippines] = &H3409 '' Windows 98/Me, Windows 2000 and later: English (Philippines)    en-PH   Latn    1252
    [Locale English/Singapore] = &H4809 '' Windows Vista and later: English (Singapore)    en-SG   Latn    1252
    [Locale English/South Africa] = &H1C09 '' English (South Africa)  en-ZA   Latn    1252
    [Locale English/Trinidad and Tobago] = &H2C09 '' English (Trinidad and Tobago)   en-TT   Latn    1252
    [Locale English/United Kingdom] = &H809  '' English (United Kingdom)    en-GB   Latn    1252
    [Locale English/United States] = &H409  '' English (United States)     en-US   Latn    1252
    [Locale English/Zimbabwe] = &H3009 '' Windows 98/Me, Windows 2000 and later: English (Zimbabwe)   en-ZW   Latn    1252
    [Locale Estonian] = &H425  '' Estonian (Estonia)  et-EE   Latn    1257
    [Locale Faroese] = &H438  '' Faroese (Faroe Islands)     fo-FO   Latn    1252
    [Locale Filipino] = &H464  '' Windows XP SP2 and later (downloadable); Windows Vista and later: Filipino (Philippines)    fil-PH  Latn    1252
    [Locale Finnish] = &H40B  '' Finnish (Finland)   fi-FI   Latn    1252
    [Locale French/Belgium] = &H80C  '' French (Belgium)    fr-BE   Latn    1252
    [Locale French/Canada] = &HC0C  '' French (Canada)     fr-CA   Latn    1252
    [Locale French/France] = &H40C  '' French (France)     fr-FR   Latn    1252
    [Locale French/Luxembourg] = &H140C '' French (Luxembourg)     fr-LU   Latn    1252
    [Locale French/Monaco] = &H180C '' French (Monaco)     fr-MC   Latn    1252
    [Locale French/Switzerland] = &H100C '' French (Switzerland)    fr-CH   Latn    1252
    [Locale Frisian] = &H462  '' Windows XP SP2 and later (downloadable); Windows Vista and later: Frisian (Netherlands)     fy-NL   Latn    1252
    [Locale Galician] = &H456  '' Windows XP and later: Galician (Spain)  gl-ES   Latn    1252
    [Locale Georgian] = &H437  '' Windows 2000 and later: Georgian (Georgia)  ka-GE   Geor    Unicode only
    [Locale German/Austria] = &HC07  '' German (Austria)    de-AT   Latn    1252
    [Locale German/Germany] = &H407  '' German (Germany)    de-DE   Latn    1252
    [Locale German/Liechtenstein] = &H1407 '' German (Liechtenstein)  de-LI   Latn    1252
    [Locale German/Luxembourg] = &H1007 '' German (Luxembourg)     de-LU   Latn    1252
    [Locale German/Switzerland] = &H807  '' German (Switzerland)    de-CH   Latn    1252
    [Locale Greek] = &H408  '' Greek (Greece)  el-GR   Grek    1253
    [Locale Greenlandic] = &H46F  '' Windows Vista and later: Greenlandic (Greenland)    kl-GL   Latn    1252
    [Locale Gujarati] = &H447  '' Windows XP and later: Gujarati (India)  gu-IN   Gujr    Unicode only
    [Locale Hausa] = &H468  '' Windows Vista and later: Hausa (Nigeria, Latin)     ha-Latn-NG  Latn    1252
    [Locale Hebrew] = &H40D  '' Hebrew (Israel)     he-IL   Hebr    1255
    [Locale Hindi] = &H439  '' Windows 2000 and later: Hindi (India)   hi-IN   Deva    Unicode only
    [Locale Hungarian] = &H40E  '' Hungarian (Hungary)     hu-HU   Latn    1250
    [Locale Icelandic] = &H40F  '' Icelandic (Iceland)     is-IS   Latn    1252
    [Locale Igbo] = &H470  '' Igbo (Nigeria)  ig-NG
    [Locale Indonesian] = &H421  '' Indonesian (Indonesia)  id-ID   Latn    1252
    [Locale Inuktitut/Latin] = &H85D  '' Windows XP and later: Inuktitut (Canada, Latin)     iu-Latn-CA  Latn    1252
    [Locale Inuktitut/Syllabics] = &H45D  '' Windows XP SP2 and later (downloadable); Windows Vista and later: Inuktitut (Canada, Syllabics)     iu-Cans-CA  Cans    Unicode only
    [Locale Irish] = &H83C  '' Windows XP SP2 and later (downloadable); Windows Vista and later: Irish (Ireland)   ga-IE   Latn    1252
    [Locale Italian/Italy] = &H410  '' Italian (Italy)     it-IT   Latn    1252
    [Locale Italian/Switzerland] = &H810  '' Italian (Switzerland)   it-CH   Latn    1252
    [Locale Japanese] = &H411  '' Japanese (Japan)    ja-JP   Hani;Hira;Kana  932
    [Locale Kannada] = &H44B  '' Windows XP and later: Kannada (India)   kn-IN   Knda    Unicode only
    [Locale Kazakh] = &H43F  '' Windows 2000 and later: Kazakh (Kazakhstan)     kk-KZ   Cyrl    1251
    [Locale Khmer] = &H453  '' Windows Vista and later: Khmer (Cambodia)   kh-KH   Khmr    Unicode only
    [Locale K'iche] = &H486  '' Windows Vista And later: K 'iche (Guatemala)     qut-GT  Latn    1252
    [Locale Kinyarwanda] = &H487  '' Windows Vista and later: Kinyarwanda (Rwanda)   rw-RW   Latn    1252
    [Locale Konkani] = &H457  '' Windows 2000 and later: Konkani (India)     kok-IN  Deva    Unicode only
    [Locale Korean/Johab] = &H812  '' Windows 95, Windows NT 4.0 only: Korean (Johab)
    [Locale Korean/Korea] = &H412  '' Korean (Korea)  ko-KR   Hang;Hani   949
    [Locale Kyrgyz] = &H440  '' Windows XP and later: Kyrgyz (Kyrgyzstan)   ky-KG   Cyrl    1251
    [Locale Lao] = &H454  '' Windows Vista and later: Lao (Lao PDR)  lo-LA   Laoo    Unicode only
    [Locale Latvian] = &H426  '' Latvian (Latvia)    lv-LV   Latn    1257
    [Locale Lithuanian] = &H427  '' Lithuanian (Lithuania)  lt-LT   Latn    1257
    [Locale Lower Sorbian] = &H82E  '' Windows Vista and later: Lower Sorbian (Germany)    dsb-DE  Latn    1252
    [Locale Luxembourgish] = &H46E  '' Windows XP SP2 and later (downloadable); Windows Vista and later: Luxembourgish (Luxembourg)    lb-LU   Latn    1252
    [Locale Macedonia] = &H42F  '' Windows 2000 and later: Macedonian (Macedonia, FYROM)   mk-MK   Cyrl    1251
    [Locale Malay/Brunei Darussalam] = &H83E  '' Windows 2000 and later: Malay (Brunei Darussalam)   ms-BN   Latn    1252
    [Locale Malay/Malaysia] = &H43E  '' Windows 2000 and later: Malay (Malaysia)    ms-MY   Latn    1252
    [Locale Malayalam] = &H44C  '' Windows XP SP2 and later: Malayalam (India)     ml-IN   Mlym    Unicode only
    [Locale Maltese] = &H43A  '' Windows XP SP2 and later: Maltese (Malta)   mt-MT   Latn    1252
    [Locale Maori] = &H481  '' Windows XP SP2 and later: Maori (New Zealand)   mi-NZ   Latn    1252
    [Locale Mapudungun] = &H47A  '' Windows XP SP2 and later (downloadable); Windows Vista and later: Mapudungun (Chile)    arn-CL  Latn    1252
    [Locale Marathi] = &H44E  '' Windows 2000 and later: Marathi (India)     mr-IN   Deva    Unicode only
    [Locale Mohawk] = &H47C  '' Windows XP SP2 and later (downloadable); Windows Vista and later: Mohawk (Canada)   moh-CA  Latn    1252
    [Locale Mongolian/Mongolia] = &H450  '' Windows XP and later: Mongolian (Mongolia)  mn-Cyrl-MN  Cyrl    1251
    [Locale Mongolian/PRC] = &H850  '' Windows Vista and later: Mongolian (PRC)    mn-Mong-CN  Mong    Unicode only
    [Locale Nepali] = &H461  '' Windows XP SP2 and later (downloadable); Windows Vista and later: Nepali (Nepal)    ne-NP   Deva    Unicode only
    [Locale Norwegian/Bokmål] = &H414  '' Norwegian (Bokmål, Norway)  nb-NO   Latn    1252
    [Locale Norwegian/Nynorsk] = &H814  '' Norwegian (Nynorsk, Norway)     nn-NO   Latn    1252
    [Locale Occitan] = &H482  '' Occitan (France)    oc-FR   Latn    1252
    [Locale Oriya] = &H448  '' Oriya (India)   or-IN   Orya    Unicode only
    [Locale Pashto] = &H463  '' Windows XP SP2 and later (downloadable); Windows Vista and later: Pashto (Afghanistan)  ps-AF
    [Locale Persian] = &H429  '' Persian (Iran)  fa-IR   Arab    1256
    [Locale Polish] = &H415  '' Polish (Poland)     pl-PL   Latn    1250
    [Locale Portuguese/Brazil] = &H416  '' Portuguese (Brazil)     pt-BR   Latn    1252
    [Locale Portuguese/Portugal] = &H816  '' Portuguese (Portugal)   pt-PT   Latn    1252
    [Locale Punjabi] = &H446  '' Windows XP and later: Punjabi (India)   pa-IN   Guru    Unicode only
    [Locale Quechua/Bolivia] = &H46B  '' Windows XP SP2 and later: Quechua (Bolivia)     quz-BO  Latn    1252
    [Locale Quechua/Ecuador] = &H86B  '' Windows XP SP2 and later: Quechua (Ecuador)     quz-EC  Latn    1252
    [Locale Quechua/Peru] = &HC6B  '' Windows XP SP2 and later: Quechua (Peru)    quz-PE  Latn    1252
    [Locale Romanian] = &H418  '' Romanian (Romania)  ro-RO   Latn    1250
    [Locale Romansh] = &H417  '' Windows XP SP2 and later (downloadable); Windows Vista and later: Romansh (Switzerland)     rm-CH   Latn    1252
    [Locale Russian] = &H419  '' Russian (Russia)    ru-RU   Cyrl    1251
    [Locale Sami/Inari] = &H243B '' Windows XP SP2 and later: Sami (Inari, Finland)     smn-FI  Latn    1252
    [Locale Sami/Lule/Norway] = &H103B '' Windows XP SP2 and later: Sami (Lule, Norway)   smj-NO  Latn    1252
    [Locale Sami/Lule/Sweden] = &H143B '' Windows XP SP2 and later: Sami (Lule, Sweden)   smj-SE  Latn    1252
    [Locale Sami/Northern/Finland] = &HC3B  '' Windows XP SP2 and later: Sami (Northern, Finland)  se-FI   Latn    1252
    [Locale Sami/Northern/Norway] = &H43B  '' Windows XP SP2 and later: Sami (Northern, Norway)   se-NO   Latn    1252
    [Locale Sami/Northern/Sweden] = &H83B  '' Windows XP SP2 and later: Sami (Northern, Sweden)   se-SE   Latn    1252
    [Locale Sami/Skolt/Finland] = &H203B '' Windows XP SP2 and later: Sami (Skolt, Finland)     sms-FI  Latn    1252
    [Locale Sami/Southern/Norway] = &H183B '' Windows XP SP2 and later: Sami (Southern, Norway)   sma-NO  Latn    1252
    [Locale Sami/Southern/Sweden] = &H1C3B '' Windows XP SP2 and later: Sami (Southern, Sweden)   sma-SE  Latn    1252
    [Locale Sanskrit] = &H44F  '' Windows 2000 and later: Sanskrit (India)    sa-IN   Deva    Unicode only
    [Locale Serbian/Bosnia and Herzegovina/Cyrillic] = &H1C1A '' Windows XP SP2 and later: Serbian (Bosnia and Herzegovina, Cyrillic)    sr-Cyrl-BA  Cyrl    1251
    [Locale Serbian/Bosnia and Herzegovina/Latin] = &H181A '' Windows XP SP2 and later: Serbian (Bosnia and Herzegovina, Latin)   sr-Latn-BA  Latn    1250
    [Locale Serbian/Serbia/Cyrillic] = &HC1A  '' Serbian (Serbia, Cyrillic)  sr-Cyrl-CS  Cyrl    1251
    [Locale Serbian/Serbia/Latin] = &H81A  '' Serbian (Serbia, Latin)     sr-Latn-CS  Latn    1250
    [Locale Sesotho sa Leboa] = &H46C  '' Windows XP SP2 and later: Sesotho sa Leboa/Northern Sotho (South Africa)    ns-ZA   Latn    1252
    [Locale Setswana / Tswana] = &H432  '' Windows XP SP2 and later: Setswana/Tswana (South Africa)    tn-ZA   Latn    1252
    [Locale Sinhala] = &H45B  '' Windows Vista and later: Sinhala (Sri Lanka)    si-LK   Sinh    Unicode only
    [Locale Slovak] = &H41B  '' Slovak (Slovakia)   sk-SK   Latn    1250
    [Locale Slovenian] = &H424  '' Slovenian (Slovenia)    sl-SI   Latn    1250
    [Locale Spanish/Argentina] = &H2C0A '' Spanish (Argentina)     es-AR   Latn    1252
    [Locale Spanish/Bolivia] = &H400A '' Spanish (Bolivia)   es-BO   Latn    1252
    [Locale Spanish/Chile] = &H340A '' Spanish (Chile)     es-CL   Latn    1252
    [Locale Spanish/Colombia] = &H240A '' Spanish (Colombia)  es-CO   Latn    1252
    [Locale Spanish/Costa Rica] = &H140A '' Spanish (Costa Rica)    es-CR   Latn    1252
    [Locale Spanish/Dominican Republic] = &H1C0A '' Spanish (Dominican Republic)    es-DO   Latn    1252
    [Locale Spanish/Ecuador] = &H300A '' Spanish (Ecuador)   es-EC   Latn    1252
    [Locale Spanish/El Salvador] = &H440A '' Spanish (El Salvador)   es-SV   Latn    1252
    [Locale Spanish/Guatemala] = &H100A '' Spanish (Guatemala)     es-GT   Latn    1252
    [Locale Spanish/Honduras] = &H480A '' Spanish (Honduras)  es-HN   Latn    1252
    [Locale Spanish/Mexico] = &H80A  '' Spanish (Mexico)    es-MX   Latn    1252
    [Locale Spanish/Nicaragua] = &H4C0A '' Spanish (Nicaragua)     es-NI   Latn    1252
    [Locale Spanish/Panama] = &H180A '' Spanish (Panama)    es-PA   Latn    1252
    [Locale Spanish/Paraguay] = &H3C0A '' Spanish (Paraguay)  es-PY   Latn    1252
    [Locale Spanish/Peru] = &H280A '' Spanish (Peru)  es-PE   Latn    1252
    [Locale Spanish/Puerto Rico] = &H500A '' Spanish (Puerto Rico)   es-PR   Latn    1252
    [Locale Spanish/Spain] = &HC0A  '' Spanish (Spain)     es-ES   Latn    1252
    [Locale Spanish/Spain Traditional] = &H40A  '' Spanish (Spain, Traditional Sort)   es-ES_tradnl    Latn    1252
    [Locale Spanish/United States] = &H540A '' Windows Vista and later: Spanish (United States)    es-US
    [Locale Spanish/Uruguay] = &H380A '' Spanish (Uruguay)   es-UY   Latn    1252
    [Locale Spanish/Venezuela] = &H200A '' Spanish (Venezuela)     es-VE   Latn    1252
    [Locale Sutu] = &H430  '' Not supported: Sutu
    [Locale Swahili] = &H441  '' Windows 2000 and later: Swahili (Kenya)     sw-KE   Latn    1252
    [Locale Swedish/Finland] = &H81D  '' Swedish (Finland)   sv-FI   Latn    1252
    [Locale Swedish/Sweden] = &H41D  '' Swedish (Sweden)    sv-SE   Latn    1252
    [Locale Syriac] = &H45A  '' Windows XP and later: Syriac (Syria)    syr-SY  Syrc    Unicode only
    [Locale Tajik] = &H428  '' Windows Vista and later: Tajik (Tajikistan)     tg-Cyrl-TJ  Cyrl    1251
    [Locale Tamazight] = &H85F  '' Windows Vista and later: Tamazight (Algeria, Latin)     tmz-Latn-DZ     Latn    1252
    [Locale Tamil] = &H449  '' Windows 2000 and later: Tamil (India)   ta-IN   Taml    Unicode only
    [Locale Tatar] = &H444  '' Windows XP and later: Tatar (Russia)    tt-RU   Cyrl    1251
    [Locale Telugu] = &H44A  '' Windows XP and later: Telugu (India)    te-IN   Telu    Unicode only
    [Locale Thai] = &H41E  '' Thai (Thailand)     th-TH   Thai    874
    [Locale Tibetan/Bhutan] = &H851  '' Windows Vista and later: Tibetan (Bhutan)   bo-BT   Tibt    Unicode only
    [Locale Tibetan/PRC] = &H451  '' Windows Vista and later: Tibetan (PRC)  bo-CN   Tibt    Unicode only
    [Locale Turkish] = &H41F  '' Turkish (Turkey)    tr-TR   Latn    1254
    [Locale Turkmen] = &H442  '' Windows Vista and later: Turkmen (Turkmenistan)     tk-TM   Cyrl    1251
    [Locale Uighur] = &H480  '' Windows Vista and later: Uighur (PRC)   ug-CN   Arab    1256
    [Locale Ukrainian] = &H422  '' Ukrainian (Ukraine)     uk-UA   Cyrl    1251
    [Locale Upper Sorbian] = &H42E  '' Windows Vista and later: Upper Sorbian (Germany)    wen-DE  Latn    1252
    [Locale Urdu/India] = &H820  '' Urdu (India)    tr-IN
    [Locale Urdu/Pakistan] = &H420  '' Windows 98/Me, Windows 2000 and later: Urdu (Pakistan)  ur-PK   Arab    1256
    [Locale Uzbek/Cyrillic] = &H843  '' Windows 2000 and later: Uzbek (Uzbekistan, Cyrillic)    uz-Cyrl-UZ  Cyrl    1251
    [Locale Uzbek/Latin] = &H443  '' Windows 2000 and later: Uzbek (Uzbekistan, Latin)   uz-Latn-UZ  Latn    1254
    [Locale Vietnamese] = &H42A  '' Windows 98/Me, Windows NT 4.0 and later: Vietnamese (Vietnam)   vi-VN   Latn    1258
    [Locale Welsh] = &H452  '' Windows XP SP2 and later: Welsh (United Kingdom)    cy-GB   Latn    1252
    [Locale Wolof] = &H488  '' Windows Vista and later: Wolof (Senegal)    wo-SN   Latn    1252
    [Locale Xhosa / isiXhosa] = &H434  '' Windows XP SP2 and later: Xhosa/isiXhosa (South Africa)     xh-ZA   Latn    1252
    [Locale Yakut] = &H485  '' Windows Vista and later: Yakut (Russia)     sah-RU  Cyrl    1251
    [Locale Yi] = &H478  '' Windows Vista and later: Yi (PRC)   ii-CN   Yiii    Unicode only
    [Locale Yoruba] = &H46A  '' Windows Vista and later: Yoruba (Nigeria)   yo-NG
    [Locale Zulu / isiZulu] = &H435  '' Windows XP SP2 and later: Zulu/isiZulu (South Africa)   zu-ZA   Latn    1252
End Enum

' Mouse button constants
Public Enum UniListMouseButton
    [No Button] = 0
    [Left Button] = vbLeftButton
    [Right Button] = vbRightButton
    [Left And Right Button] = vbLeftButton Or vbRightButton
    [Middle Button] = vbMiddleButton
    [Left And Middle Button] = vbLeftButton Or vbMiddleButton
    [Right And Middle Button] = vbRightButton Or vbMiddleButton
    [All Buttons] = vbLeftButton Or vbRightButton Or vbMiddleButton
End Enum

Public Enum UniListMouseWheel
    [Wheel Down]
    [Wheel Up]
End Enum

Public Enum UniListScrollBarVisibility
    [Allow No ScrollBars] = False
    [Disable No ScrollBars] = True
End Enum

' Scroll direction constants
Public Enum UniListScrollDirection
    [Scroll Horizontal] = vbHorizontal
    [Scroll Vertical] = vbVertical
End Enum

' Shift constants
Public Enum UniListShift
    [No Mask] = 0
    [Shift Mask] = vbShiftMask
    [Ctrl Mask] = vbCtrlMask
    [Shift And Ctrl Mask] = vbShiftMask Or vbCtrlMask
    [Alt Mask] = vbAltMask
    [Shift And Alt Mask] = vbShiftMask Or vbAltMask
    [Ctrl And Alt Mask] = vbCtrlMask Or vbAltMask
    [All Masks] = vbShiftMask Or vbCtrlMask Or vbAltMask
End Enum

Private Const BDR_INNER = &HC
Private Const BDR_OUTER = &H3
Private Const BDR_RAISED = &H5
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKEN = &HA
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2
Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Private Const LB_ADDSTRING As Long = &H180
Private Const LB_RESETCONTENT As Long = &H184&
Private Const LB_DELETESTRING As Long = &H182&
Private Const LB_INSERTSTRING As Long = &H181&
Private Const LB_FINDSTRING As Long = &H18F&
Private Const LB_FINDSTRINGEXACT As Long = &H1A2&
Private Const LB_SELECTSTRING As Long = &H18C&
Private Const LB_GETTEXT As Long = &H189&
Private Const LB_GETTEXTLEN As Long = &H18A&
Private Const LB_SETTEXT As Long = &H1AA&
Private Const LB_GETITEMDATA As Long = &H199&
Private Const LB_SETITEMDATA As Long = &H19A&
Private Const LB_SETTABSTOPS As Long = &H192&
Private Const LB_GETITEMHEIGHT As Long = &H1A1&
Private Const LB_GETITEMRECT As Long = &H198&
Private Const LB_GETTOPINDEX As Long = &H18E&
Private Const LB_SETTOPINDEX = &H197
Private Const LB_SELITEMRANGEEX As Long = &H183&
Private Const LB_SELITEMRANGE As Long = &H19B&
Private Const LB_GETCOUNT As Long = &H18B&
Private Const LB_GETCURSEL As Long = &H188&
Private Const LB_SETCURSEL As Long = &H186&
Private Const LB_GETSELCOUNT As Long = &H190&
Private Const LB_GETSELITEMS As Long = &H191&
Private Const LB_GETCARETINDEX As Long = &H19F&
Private Const LB_SETCARETINDEX As Long = &H19E&
Private Const LB_GETSEL As Long = &H187&
Private Const LB_SETSEL As Long = &H185&
Private Const LB_GETLOCALE = &H1A6
Private Const LB_SETLOCALE = &H1A5
Private Const LB_SETCOLUMNWIDTH = &H195
Private Const LB_INITSTORAGE = &H1A8
Private Const LB_GETANCHORINDEX = &H19D
Private Const LB_SETANCHORINDEX = &H19C
Private Const LB_GETHORIZONTALEXTENT = &H193
Private Const LB_SETHORIZONTALEXTENT = &H194

Private Const LBS_NOTIFY = &H1
Private Const LBS_SORT = &H2
Private Const LBS_NOREDRAW = &H4
Private Const LBS_MULTIPLESEL = &H8
Private Const LBS_OWNERDRAWFIXED = &H10
Private Const LBS_OWNERDRAWVARIABLE = &H20
Private Const LBS_HASSTRINGS = &H40
Private Const LBS_USETABSTOPS = &H80
Private Const LBS_NOINTEGRALHEIGHT = &H100
Private Const LBS_MULTICOLUMN = &H200
Private Const LBS_WANTKEYBOARDINPUT = &H400
Private Const LBS_EXTENDEDSEL = &H800
Private Const LBS_DISABLENOSCROLL = &H1000
Private Const LBS_NODATA = &H2000
Private Const LBS_NOSEL = &H4000

Private Const ES_AUTOVSCROLL = &H40&
Private Const ES_AUTOHSCROLL = &H80&
Private Const ES_CENTER = &H1&
Private Const ES_LEFT = &H0&
Private Const ES_NOHIDESEL = &H100&
Private Const ES_RIGHT = &H2&
Private Const ES_WANTRETURN = &H1000&

Private Const GWL_EXSTYLE As Long = -20&
Private Const GWL_STYLE As Long = -16&

Private Const VK_DOWN = &H28
Private Const VK_END = &H23
Private Const VK_HOME = &H24
Private Const VK_LEFT = &H25
Private Const VK_RETURN = &HD
Private Const VK_RIGHT = &H27
Private Const VK_TAB = &H9
Private Const VK_UP = &H26

Private Const WS_BORDER = &H800000
Private Const WS_CHILD = &H40000000
Private Const WS_HSCROLL = &H100000
Private Const WS_VISIBLE = &H10000000
Private Const WS_VSCROLL = &H200000

' WM_MOUSEMOVE and others
Private Const MK_LBUTTON = &H1&
Private Const MK_RBUTTON = &H2&
Private Const MK_SHIFT = &H4&
Private Const MK_CONTROL = &H8&
Private Const MK_MBUTTON = &H10&

' window messages
Private Const WM_CHAR = &H102&
Private Const WM_COMMAND = &H111&
Private Const WM_CREATE = &H1&
Private Const WM_CTLCOLOREDIT = &H133&
Private Const WM_CTLCOLORSTATIC = &H138&
Private Const WM_DESTROY = &H2&
Private Const WM_ERASEBKGND = &H14
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_HSCROLL = &H114
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_KILLFOCUS = &H8
Private Const WM_LBUTTONDBLCLK = &H203&
Private Const WM_LBUTTONDOWN = &H201&
Private Const WM_LBUTTONUP = &H202&
Private Const WM_MBUTTONDBLCLK = &H209&
Private Const WM_MBUTTONDOWN = &H207&
Private Const WM_MBUTTONUP = &H208&
Private Const WM_MOUSEACTIVATE = &H21
Private Const WM_MOUSEFIRST = &H200
Private Const WM_MOUSELAST = &H209
Private Const WM_MOUSEMOVE = &H200&
Private Const WM_MOUSELEAVE = &H2A3&
Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_PAINT = &HF&
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204&
Private Const WM_RBUTTONUP = &H205&
Private Const WM_SETFOCUS = &H7
Private Const WM_SETFONT = &H30
Private Const WM_SETTEXT = &HC
Private Const WM_UNDO = &H304
Private Const WM_USER = &H400
Private Const WM_VSCROLL = &H115

Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_LAYOUTRTL = &H400000

Private Const OPAQUE = 2
Private Const TRANSPARENT = 1

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize As Long
    dwFlags As TRACKMOUSEEVENT_FLAGS
    hwndTrack As Long
    dwHoverTime As Long
End Type

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Const SB_THUMBPOSITION = 4
Private Const SBS_HORZ = &H0&
Private Const SBS_VERT = &H1&

Private Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long
Private Declare Function PostMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function TrackMouseEventUser32 Lib "user32" Alias "TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetACP Lib "kernel32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, lpUsedDefaultChar As Long) As Long

Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function SetKeyboardState Lib "user32" (pbKeyState As Byte) As Long

Private Declare Function GetDialogBaseUnits Lib "user32" () As Long

Private Declare Function LocalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LocalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Private m_Appearance As UniListAppearance
Private m_BorderStyle As UniListBorderStyle
Private m_CaptureEnter As Boolean
Private m_CaptureEsc As Boolean
Private m_CaptureTab As Boolean
Private m_Columns As Byte
Private m_DisableSelect As Boolean
Private m_Enabled As Boolean
Private WithEvents m_Font As StdFont
Attribute m_Font.VB_VarHelpID = -1
Private m_ForeColor As OLE_COLOR
Private m_HasFocus As Boolean
Private m_IntegralHeight As Boolean
Private m_MouseOver As Boolean
Private m_MultiSelect As MultiSelectConstants
Private m_NewIndex As Long
Private m_RightToLeft As Boolean
Private m_ScrollBars As ScrollBarConstants
Private m_ScrollBarVisibility As UniListScrollBarVisibility
Private m_ScrollWidth As Long
Private m_Sort As Boolean
Private m_StorageItems As Long
Private m_StorageMB As Single
Private m_Style As ListBoxConstants
Private m_UseEvents As Boolean
Private m_UseTabStops As Boolean

Private m_BackClr As Long
Private m_BackClrBrush As Long
Private m_Focus As Boolean
Private m_ForeClr As Long
Private m_hDC As Long
Private m_hWnd As Long
Private m_IPAO As UniList_IPAOHook
Private m_RC As RECT
Private m_ContainerScaleMode As ScaleModeConstants   ' Container ScaleMode
Private m_TrackComCtl As Boolean
Private m_TrackUser32 As Boolean

' for fixing XP Theme problem with a certain version of comctl32.dll
Private m_FreeShell32 As Boolean
Private m_Shell32 As Long

Dim blnDesignTime As Boolean                ' True if in IDE design time

    Private z_scFunk            As Collection   'hWnd/thunk-address collection; initialized as needed
    Private z_hkFunk            As Collection   'hook/thunk-address collection; initialized as needed
    Private z_cbFunk            As Collection   'callback/thunk-address collection; initialized as needed
    Private Const IDX_INDEX     As Long = 2     'index of the subclassed hWnd OR hook type
    Private Const IDX_PREVPROC  As Long = 9     'Thunk data index of the original WndProc
    Private Const IDX_BTABLE    As Long = 11    'Thunk data index of the Before table for messages
    Private Const IDX_ATABLE    As Long = 12    'Thunk data index of the After table for messages
    Private Const IDX_CALLBACKORDINAL As Long = 36 ' Ubound(callback thunkdata)+1, index of the callback

  ' Declarations:
    Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
    Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
    Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
    Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
    Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
    Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    Private Enum eThunkType
        SubclassThunk = 0
        HookThunk = 1
        CallbackThunk = 2
    End Enum

    Private Enum eMsgWhen                                                   'When to callback
      MSG_BEFORE = 1                                                        'Callback before the original WndProc
      MSG_AFTER = 2                                                         'Callback after the original WndProc
      MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                            'Callback before and after the original WndProc
    End Enum
    
    ' see ssc_Subclass for complete listing of indexes and what they relate to
    Private Const IDX_PARM_USER As Long = 13    'Thunk data index of the User-defined callback parameter data index
    Private Const IDX_UNICODE   As Long = 107   'Must be UBound(subclass thunkdata)+1; index for unicode support
    Private Const MSG_ENTRIES   As Long = 32    'Number of msg table entries. Set to 1 if using ALL_MESSAGES for all subclassed windows
    Private Const ALL_MESSAGES  As Long = -1    'All messages will callback
    
    Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function CallWindowProcW Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
    Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hWnd As Long) As Long
    Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

    '-SelfHook specific declarations----------------------------------------------------------------------------
    Private Declare Function SetWindowsHookExA Lib "user32" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function SetWindowsHookExW Lib "user32" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function UnHookWindowsHookEx Lib "user32" Alias "UnhookWindowsHookEx" (ByVal hHook As Long) As Long
    Private Declare Function CallNextHookEx Lib "user32.dll" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
    Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
    
    Private Enum eHookType  ' http://msdn2.microsoft.com/en-us/library/ms644990.aspx
      WH_MSGFILTER = -1
      WH_JOURNALRECORD = 0
      WH_JOURNALPLAYBACK = 1
      WH_KEYBOARD = 2
      WH_GETMESSAGE = 3
      WH_CALLWNDPROC = 4
      WH_CBT = 5
      WH_SYSMSGFILTER = 6
      WH_MOUSE = 7
      WH_DEBUG = 9
      WH_SHELL = 10
      WH_FOREGROUNDIDLE = 11
      WH_CALLWNDPROCRET = 12
      WH_KEYBOARD_LL = 13       ' NT/2000/XP+ only, Global hook only
      WH_MOUSE_LL = 14          ' NT/2000/XP+ only, Global hook only
    End Enum

Public Function AddItem(ByRef Text As String, Optional ByVal BeforeIndex As Long = -1&, Optional ByVal ItemData As Long = 0&, Optional ByVal Checked As Boolean) As Long
    If Not m_Sort Then
        m_NewIndex = SendMessageW(m_hWnd, LB_INSERTSTRING, BeforeIndex, ByVal StrPtr(Text))
    ElseIf (BeforeIndex > -1&) Then
        m_NewIndex = SendMessageW(m_hWnd, LB_INSERTSTRING, BeforeIndex, ByVal StrPtr(Text))
    ElseIf BeforeIndex = -1& Then
        m_NewIndex = SendMessageW(m_hWnd, LB_ADDSTRING, 0&, ByVal StrPtr(Text))
    Else
        AddItem = -1&
        Exit Function
    End If
    If ItemData <> 0 Then SendMessageW m_hWnd, LB_SETITEMDATA, m_NewIndex, ByVal ItemData
    If Checked Then
        If m_MultiSelect = vbMultiSelectNone Then
            Debug.Print SendMessageW(m_hWnd, LB_SETCARETINDEX, NewIndex, ByVal 0&)
        Else
            If SendMessageW(m_hWnd, LB_SETSEL, 1, ByVal m_NewIndex) > -1 Then
                If m_Style = vbListBoxCheckbox Then RaiseEvent ItemCheck(m_NewIndex)
            End If
        End If
    End If
    AddItem = m_NewIndex
End Function
Public Property Get Appearance() As UniListAppearance
    Appearance = m_Appearance
End Property
Public Property Let Appearance(ByVal newValue As UniListAppearance)
    m_Appearance = newValue
    Private_Init
End Property
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal NewColor As OLE_COLOR)
    UserControl.BackColor = NewColor
    UserControl.Refresh
    If NewColor < 0 Then m_BackClr = GetSysColor(NewColor And &HFF&) Else m_BackClr = NewColor
    DeleteObject m_BackClrBrush
    m_BackClrBrush = CreateSolidBrush(m_BackClr)
    If Not blnDesignTime Then Else PropertyChanged "BackColor"
    InvalidateRect m_hWnd, m_RC, -1&
End Property
Public Property Get BorderStyle() As UniListBorderStyle
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BorderStyle.VB_UserMemId = -504
    BorderStyle = m_BorderStyle
End Property
Public Property Let BorderStyle(ByVal newValue As UniListBorderStyle)
    m_BorderStyle = newValue
    If Not blnDesignTime Then Else PropertyChanged "BorderStyle"
    If m_Appearance = [Classic 3D] Then
        UserControl_Resize
    Else
        Private_Init
    End If
End Property
Public Property Get CaptureEnter() As Boolean
    CaptureEnter = m_CaptureEnter
End Property
Public Property Let CaptureEnter(ByVal newValue As Boolean)
    m_CaptureEnter = newValue
    If Not blnDesignTime Then Else PropertyChanged "CaptureEnter"
End Property
Public Property Get CaptureEsc() As Boolean
    CaptureEsc = m_CaptureEsc
End Property
Public Property Let CaptureEsc(ByVal newValue As Boolean)
    m_CaptureEsc = newValue
    If Not blnDesignTime Then Else PropertyChanged "CaptureEsc"
End Property
Public Property Get CaptureTab() As Boolean
    CaptureTab = m_CaptureTab
End Property
Public Property Let CaptureTab(ByVal newValue As Boolean)
    m_CaptureTab = newValue
    If Not blnDesignTime Then Else PropertyChanged "CaptureTab"
End Property
Public Property Get Caret() As Long
Attribute Caret.VB_MemberFlags = "400"
    ListIndex = SendMessageW(m_hWnd, LB_GETCARETINDEX, 0&, ByVal 0&)
End Property
Public Property Let Caret(ByVal NewIndex As Long)
    SendMessageW m_hWnd, LB_SETCARETINDEX, NewIndex, ByVal 0&
End Property
Public Sub Clear()
    If m_hWnd Then SendMessageW m_hWnd, LB_RESETCONTENT, 0&, ByVal 0&
End Sub
Public Property Get Columns() As Byte
    Columns = m_Columns
End Property
Public Property Let Columns(ByVal newValue As Byte)
    If m_Columns <> newValue Then
        m_Columns = newValue
        Private_Init
        If Not blnDesignTime Then Else PropertyChanged "Columns"
    End If
End Property
Public Property Get DisableSelect() As Boolean
    DisableSelect = m_DisableSelect
End Property
Public Property Let DisableSelect(ByVal newValue As Boolean)
    If m_DisableSelect <> newValue Then
        m_DisableSelect = newValue
        Private_Init
        If Not blnDesignTime Then Else PropertyChanged "DisableSelect"
    End If
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = m_Enabled
End Property
Public Property Let Enabled(ByVal newValue As Boolean)
    Dim lngStyle As Long
    If (m_Enabled <> newValue) And (m_hWnd <> 0) Then
        UserControl.Enabled = newValue
        m_Enabled = newValue
        If Not blnDesignTime Then Else PropertyChanged "Enabled"
    End If
End Property
Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property
Public Property Set Font(ByVal newValue As Font)
    Dim NewFont As New StdFont
    ' have to do it this way because otherwise we'd link with existing font object
    NewFont.Bold = newValue.Bold
    NewFont.Charset = newValue.Charset
    NewFont.Italic = newValue.Italic
    NewFont.Name = newValue.Name
    NewFont.SIZE = newValue.SIZE
    NewFont.Strikethrough = newValue.Strikethrough
    NewFont.Underline = newValue.Underline
    NewFont.Weight = newValue.Weight
    Set m_Font = NewFont
    If Not blnDesignTime Then Else PropertyChanged "Font"
    m_Font_FontChanged vbNullString
    InvalidateRect m_hWnd, m_RC, -1&
End Property
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_MemberFlags = "400"
    FontBold = m_Font.Bold
End Property
Public Property Let FontBold(ByVal newValue As Boolean)
    m_Font.Bold = newValue
    If Not blnDesignTime Then Else PropertyChanged "Font"
    InvalidateRect m_hWnd, m_RC, -1&
End Property
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_MemberFlags = "400"
    FontItalic = m_Font.Italic
End Property
Public Property Let FontItalic(ByVal newValue As Boolean)
    m_Font.Italic = newValue
    If Not blnDesignTime Then Else PropertyChanged "Font"
    InvalidateRect m_hWnd, m_RC, -1&
End Property
Public Property Get FontName() As String
Attribute FontName.VB_MemberFlags = "400"
    FontName = m_Font.Name
End Property
Public Property Let FontName(ByRef newValue As String)
    m_Font.Name = newValue
    If Not blnDesignTime Then Else PropertyChanged "Font"
    InvalidateRect m_hWnd, m_RC, -1&
End Property
Public Property Get FontSize() As Single
Attribute FontSize.VB_MemberFlags = "400"
    FontSize = m_Font.SIZE
End Property
Public Property Let FontSize(ByVal newValue As Single)
    m_Font.SIZE = newValue
    If Not blnDesignTime Then Else PropertyChanged "Font"
    InvalidateRect m_hWnd, m_RC, -1&
End Property
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_MemberFlags = "400"
    FontStrikethru = m_Font.Strikethrough
End Property
Public Property Let FontStrikethru(ByVal newValue As Boolean)
    m_Font.Strikethrough = newValue
    If Not blnDesignTime Then Else PropertyChanged "Font"
    InvalidateRect m_hWnd, m_RC, -1&
End Property
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_MemberFlags = "400"
    FontUnderline = m_Font.Underline
End Property
Public Property Let FontUnderline(ByVal newValue As Boolean)
    m_Font.Underline = newValue
    If Not blnDesignTime Then Else PropertyChanged "Font"
    InvalidateRect m_hWnd, m_RC, -1&
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal newValue As OLE_COLOR)
    m_ForeColor = newValue
    If m_ForeColor < 0 Then m_ForeClr = GetSysColor(m_ForeColor And &HFF&) Else m_ForeClr = m_ForeColor
    If Not blnDesignTime Then Else PropertyChanged "ForeColor"
    InvalidateRect m_hWnd, m_RC, -1&
End Property
Public Function GetTabStopTwip() As Long
    GetTabStopTwip = GetDialogBaseUnits
End Function
Public Property Get hdc() As Long
Attribute hdc.VB_MemberFlags = "400"
    hdc = m_hDC
End Property
Public Property Get hWnd() As Long
Attribute hWnd.VB_MemberFlags = "400"
    hWnd = m_hWnd
End Property
Public Property Get IntegralHeight() As Boolean
    IntegralHeight = m_IntegralHeight
End Property
Public Property Let IntegralHeight(ByVal newValue As Boolean)
    If m_IntegralHeight <> newValue Then
        m_IntegralHeight = newValue
        Private_Init
        If Not blnDesignTime Then Else PropertyChanged "IntegralHeight"
    End If
End Property
Public Property Get ItemData(ByVal Index As Long) As Long
    ItemData = SendMessageW(m_hWnd, LB_GETITEMDATA, Index, ByVal 0&)
End Property
Public Property Let ItemData(ByVal Index As Long, ByVal NewData As Long)
    SendMessageW m_hWnd, LB_SETITEMDATA, Index, ByVal NewData
End Property
Public Function ItemHeight() As Long
    If m_hWnd Then
        ItemHeight = SendMessageW(m_hWnd, LB_GETITEMHEIGHT, 0&, ByVal 0&)
    End If
End Function
Public Property Get List(ByVal Index As Long) As String
Attribute List.VB_MemberFlags = "200"
    Dim bytTemp() As Byte, lngLen As Long
    If m_hWnd Then
        lngLen = SendMessageW(m_hWnd, LB_GETTEXTLEN, Index, ByVal 0&)
        If lngLen > 0 Then
            ReDim bytTemp(lngLen * 2 - 1)
            If SendMessageW(m_hWnd, LB_GETTEXT, Index, ByVal VarPtr(bytTemp(0))) Then List = CStr(bytTemp)
        End If
    End If
End Property
Public Property Let List(ByVal Index As Long, ByRef newValue As String)
    If m_hWnd Then
        If LenB(newValue) Then
            SendMessageW m_hWnd, LB_SETTEXT, Index, ByVal StrPtr(newValue)
        Else
            SendMessageW m_hWnd, LB_SETTEXT, Index, ByVal 0&
        End If
    End If
End Property
Public Property Get ListCount() As Long
Attribute ListCount.VB_MemberFlags = "400"
    ListCount = SendMessageW(m_hWnd, LB_GETCOUNT, 0&, ByVal 0&)
End Property
Public Property Get ListIndex() As Long
Attribute ListIndex.VB_MemberFlags = "400"
    ListIndex = SendMessageW(m_hWnd, LB_GETCURSEL, 0&, ByVal 0&)
End Property
Public Property Let ListIndex(ByVal NewIndex As Long)
    SendMessageW m_hWnd, LB_SETCURSEL, NewIndex, ByVal 0&
End Property
Public Property Get MouseIcon() As IPictureDisp
    Set MouseIcon = UserControl.MouseIcon
End Property
Public Property Set MouseIcon(ByRef newValue As IPictureDisp)
    Set UserControl.MouseIcon = newValue
    If Not blnDesignTime Then Else PropertyChanged "MouseIcon"
End Property
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal newValue As MousePointerConstants)
    UserControl.MousePointer = newValue
    If Not blnDesignTime Then Else PropertyChanged "MousePointer"
End Property
Public Function MouseOver() As Boolean
    MouseOver = m_MouseOver
End Function
Public Property Get MultiSelect() As MultiSelectConstants
    MultiSelect = m_MultiSelect
End Property
Public Property Let MultiSelect(ByVal newValue As MultiSelectConstants)
    If m_MultiSelect <> newValue Then
        m_MultiSelect = newValue
        Private_Init
        If Not blnDesignTime Then Else PropertyChanged "MultiSelect"
    End If
End Property
Public Function NewIndex() As Long
    NewIndex = m_NewIndex
End Function
Private Function Private_GetContainerScaleMode() As ScaleModeConstants
    ' this should be called only when we know scalemode has changed
    Select Case Ambient.ScaleUnits
        Case "Twip"
            Private_GetContainerScaleMode = vbTwips
        Case "Point"
            Private_GetContainerScaleMode = vbPoints
        Case "Pixel"
            Private_GetContainerScaleMode = vbPixels
        Case "Character"
            Private_GetContainerScaleMode = vbCharacters
        Case "Inch"
            Private_GetContainerScaleMode = vbInches
        Case "Millimeter"
            Private_GetContainerScaleMode = vbMillimeters
        Case "Centimeter"
            Private_GetContainerScaleMode = vbCentimeters
        Case "User"
            ' prevent user scalemode
            UserControl.Extender.Container.ScaleMode = vbTwips
            Private_GetContainerScaleMode = vbTwips
    End Select
End Function
Private Function Private_GetShiftMask() As Integer
    Private_GetShiftMask = Abs((GetKeyState(vbKeyShift) And &H8000) = &H8000) * vbShiftMask _
        Or Abs((GetKeyState(vbKeyControl) And &H8000) = &H8000) * vbCtrlMask _
        Or Abs((GetKeyState(vbKeyMenu) And &H8000) = &H8000) * vbAltMask
End Function
Private Function Private_GetShiftState() As Long
    Private_GetShiftState = Abs((GetKeyState(vbKeyShift) And &H8000) = &H8000) * vbShiftMask _
        Or Abs((GetKeyState(vbKeyControl) And &H8000) = &H8000) * vbCtrlMask _
        Or Abs((GetKeyState(vbKeyMenu) And &H8000) = &H8000) * vbAltMask
End Function
Private Sub Private_Init()
    Dim lngStyle As Long, lngExStyle As Long, lngBytes As Long
    ' remove old
    If m_hWnd Then
        ssc_UnSubclass m_hWnd
        DestroyWindow m_hWnd
    End If
    ' init styles
    lngStyle = WS_CHILD Or WS_VISIBLE Or LBS_HASSTRINGS
    'If Not m_HideSelection Then lngStyle = lngStyle Or ES_NOHIDESEL
    If m_Columns Then
        lngStyle = lngStyle Or LBS_MULTICOLUMN
    End If
    If m_DisableSelect Then
        lngStyle = lngStyle Or LBS_NOSEL
    Else
        Select Case m_MultiSelect
            Case vbMultiSelectNone
            Case vbMultiSelectSimple
                lngStyle = lngStyle Or LBS_MULTIPLESEL
            Case vbMultiSelectExtended
                lngStyle = lngStyle Or LBS_EXTENDEDSEL
        End Select
    End If
    If Not m_IntegralHeight Then lngStyle = lngStyle Or LBS_NOINTEGRALHEIGHT
    Select Case m_ScrollBars
        Case vbSBNone
            ' ignore
        Case vbHorizontal
            lngStyle = lngStyle Or WS_HSCROLL
        Case vbVertical
            lngStyle = lngStyle Or WS_VSCROLL
        Case vbBoth
            lngStyle = lngStyle Or WS_HSCROLL Or WS_VSCROLL
    End Select
    If m_ScrollBarVisibility = [Disable No ScrollBars] Then lngStyle = lngStyle Or LBS_DISABLENOSCROLL
    If m_Sort Then lngStyle = lngStyle Or LBS_SORT
    If m_UseTabStops Then lngStyle = lngStyle Or LBS_USETABSTOPS
    
    If m_RightToLeft Then lngExStyle = WS_EX_LAYOUTRTL
    
    ' create new listbox
    If m_Appearance = [Classic 3D] Then
        m_hWnd = CreateWindowExW(lngExStyle, StrPtr("ListBox"), 0&, _
            lngStyle, _
            m_BorderStyle, m_BorderStyle, m_RC.Right - m_BorderStyle * 2, m_RC.Bottom - m_BorderStyle * 2, _
            UserControl.hWnd, 0&, App.hInstance, 0&)
    Else
        Select Case m_BorderStyle
            Case [Flat3D]
                lngStyle = lngStyle Or WS_BORDER
            Case [3D]
                lngExStyle = lngExStyle Or WS_EX_CLIENTEDGE
        End Select
        m_hWnd = CreateWindowExW(lngExStyle, StrPtr("ListBox"), 0&, _
            lngStyle, _
            0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, _
            UserControl.hWnd, 0&, App.hInstance, 0&)
    End If
    'SetWindowLong UserControl.hWnd, GWL_EXSTYLE, WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_WANTRETURN
    If m_hWnd Then
        ' get device context
        m_hDC = GetDC(m_hWnd)
        ' set font
        m_Font_FontChanged vbNullString
        ' init storage
        lngBytes = (1048576 * CDbl(m_StorageMB)) And &H7FFFFFFF
        SendMessageW m_hWnd, LB_INITSTORAGE, m_StorageItems, ByVal lngBytes
        If m_ScrollWidth > 0 And (m_ScrollBars = vbHorizontal Or m_ScrollBars = vbBoth) And m_Columns = 0 Then
            SendMessageW m_hWnd, LB_SETHORIZONTALEXTENT, m_ScrollWidth, ByVal 0&
        End If
        ' start subclassing
        If ssc_Subclass(m_hWnd, , 2, , Not blnDesignTime, True) Then
            ssc_AddMsg m_hWnd, MSG_BEFORE, _
                WM_CHAR, _
                WM_LBUTTONDOWN, _
                WM_MBUTTONDOWN, _
                WM_RBUTTONDOWN, _
                WM_MOUSEACTIVATE, _
                WM_KEYDOWN
            ssc_AddMsg m_hWnd, MSG_AFTER, _
                WM_KEYUP, _
                WM_MOUSELEAVE, _
                WM_MOUSEMOVE, _
                WM_MOUSEWHEEL, _
                WM_LBUTTONUP, _
                WM_MBUTTONUP, _
                WM_RBUTTONUP, _
                WM_LBUTTONDBLCLK, _
                WM_MBUTTONDBLCLK, _
                WM_RBUTTONDBLCLK, _
                WM_HSCROLL, _
                WM_VSCROLL, _
                WM_SETFOCUS
        End If
    Else
        m_hDC = 0
    End If
End Sub
Private Function Private_IsFunctionSupported(ByRef FunctionName As String, ByRef ModuleName As String) As Boolean
    Dim lngModule As Long, blnUnload As Boolean
    ' get handle to module
    lngModule = GetModuleHandleA(ModuleName)
    ' if getting the handle failed...
    If lngModule = 0 Then
        ' try loading the module
        lngModule = LoadLibraryA(ModuleName)
        ' we have to unload it too if that succeeded
        blnUnload = (lngModule <> 0)
    End If
    ' now if we have a handle to module...
    If lngModule Then
        ' see if the queried function is supported; return True if so, False if not
        Private_IsFunctionSupported = (GetProcAddress(lngModule, FunctionName) <> 0)
        ' see if we have to unload the module
        If blnUnload Then FreeLibrary lngModule
    End If
End Function
Private Sub Private_SetIPAO()
    Dim pOleObject As IOleObject
    Dim pOleInPlaceSite As IOleInPlaceSite
    Dim pOleInPlaceFrame As IOleInPlaceFrame
    Dim pOleInPlaceUIWindow As IOleInPlaceUIWindow
    Dim rcPos As RECT
    Dim rcClip As RECT
    Dim uFrameInfo As OLEINPLACEFRAMEINFO

    On Error Resume Next
    Set pOleObject = Me
    Set pOleInPlaceSite = pOleObject.GetClientSite
    If Not pOleInPlaceSite Is Nothing Then
        pOleInPlaceSite.GetWindowContext pOleInPlaceFrame, pOleInPlaceUIWindow, VarPtr(rcPos), VarPtr(rcClip), VarPtr(uFrameInfo)
        If Not pOleInPlaceFrame Is Nothing Then
            pOleInPlaceFrame.SetActiveObject m_IPAO.ThisPointer, vbNullString
        End If
        If Not pOleInPlaceUIWindow Is Nothing Then '-- And Not m_bMouseActivate
            pOleInPlaceUIWindow.SetActiveObject m_IPAO.ThisPointer, vbNullString
        Else
            pOleObject.DoVerb OLEIVERB_UIACTIVATE, 0, pOleInPlaceSite, 0, UserControl.hWnd, VarPtr(rcPos)
        End If
    End If
End Sub
Private Function Private_UTF16toUTF8(ByRef Text As String, Optional lFlags As Long) As String
    Static tmpArr() As Byte
    Dim tmpLen As Long, textLen As Long
    If LenB(Text) Then
        textLen = Len(Text)
        tmpLen = LenB(Text) * 2 + 1
        ReDim Preserve tmpArr(tmpLen - 1)
        tmpLen = WideCharToMultiByte(65001, lFlags, ByVal StrPtr(Text), textLen, ByVal VarPtr(tmpArr(0)), tmpLen, ByVal 0&, ByVal 0&)
        If tmpLen > 0 Then
            If UBound(tmpArr) <> tmpLen Then ReDim Preserve tmpArr(tmpLen - 1)
            Private_UTF16toUTF8 = CStr(tmpArr)
        End If
    End If
End Function
Private Function Private_UTF8toUTF16(ByRef Text As String, Optional lFlags As Long) As String
    Static tmpArr() As Byte
    Dim tmpLen As Long, textLen As Long
    If LenB(Text) Then
        textLen = LenB(Text)
        tmpLen = textLen * 2
        ReDim Preserve tmpArr(tmpLen + 1)
        tmpLen = MultiByteToWideChar(65001, lFlags, ByVal StrPtr(Text), textLen, ByVal VarPtr(tmpArr(0)), tmpLen) * 2
        If tmpLen > 0 Then
            If UBound(tmpArr) <> tmpLen Then ReDim Preserve tmpArr(tmpLen - 1)
            Private_UTF8toUTF16 = CStr(tmpArr)
        End If
    End If
End Function
Public Function RemoveItem(Index As Long) As Boolean
    If m_hWnd Then
        RemoveItem = SendMessageW(m_hWnd, LB_DELETESTRING, Index, ByVal 0&) > -1
    End If
End Function
Public Property Get RightToLeft() As Boolean
    RightToLeft = m_RightToLeft
End Property
Public Property Let RightToLeft(ByVal newValue As Boolean)
    If newValue <> m_RightToLeft Then
        m_RightToLeft = newValue
        Private_Init
        If Not blnDesignTime Then Else PropertyChanged "RightToLeft"
    End If
End Property
Public Property Get ScrollBars() As ScrollBarConstants
    ScrollBars = m_ScrollBars
End Property
Public Property Let ScrollBars(ByVal newValue As ScrollBarConstants)
    ' only change if really changed
    If m_ScrollBars <> newValue Then
        m_ScrollBars = newValue
        Private_Init
    End If
End Property
Public Property Get ScrollBarVisibility() As UniListScrollBarVisibility
    ScrollBarVisibility = m_ScrollBarVisibility
End Property
Public Property Let ScrollBarVisibility(ByVal newValue As UniListScrollBarVisibility)
    m_ScrollBarVisibility = newValue
    Private_Init
    If Not blnDesignTime Then Else PropertyChanged "ScrollBarVisibility"
End Property
Public Property Get ScrollHorizontal() As Long
Attribute ScrollHorizontal.VB_MemberFlags = "400"
    If m_hWnd <> 0 And (m_ScrollBars = vbHorizontal Or m_ScrollBars = vbBoth) Then
        ScrollHorizontal = GetScrollPos(m_hWnd, SBS_HORZ)
    End If
End Property
Public Property Let ScrollHorizontal(ByVal newValue As Long)
    If m_hWnd <> 0 And (m_ScrollBars = vbHorizontal Or m_ScrollBars = vbBoth) Then
        If newValue >= 0& And newValue <= &H7FFF& Then
            If SetScrollPos(m_hWnd, SBS_HORZ, newValue, -1&) <> -1& Then
                PostMessageW m_hWnd, WM_HSCROLL, SB_THUMBPOSITION Or (&H10000 * newValue), 0&
            End If
        End If
    End If
End Property
Public Property Get ScrollVertical() As Long
Attribute ScrollVertical.VB_MemberFlags = "400"
    If m_hWnd <> 0 And (m_ScrollBars = vbVertical Or m_ScrollBars = vbBoth) Then
        ScrollVertical = GetScrollPos(m_hWnd, SBS_VERT)
    End If
End Property
Public Property Let ScrollVertical(ByVal newValue As Long)
    If m_hWnd <> 0 And (m_ScrollBars = vbVertical Or m_ScrollBars = vbBoth) Then
        If newValue >= 0& And newValue <= &H7FFF& Then
            If SetScrollPos(m_hWnd, SBS_VERT, newValue, -1&) <> -1& Then
                PostMessageW m_hWnd, WM_VSCROLL, SB_THUMBPOSITION Or (&H10000 * newValue), 0&
            End If
        End If
    End If
End Property
Public Property Get ScrollWidth() As Long
    If m_hWnd <> 0 And (m_ScrollBars = vbHorizontal Or m_ScrollBars = vbBoth) And m_Columns = 0 Then
        ScrollWidth = SendMessageW(m_hWnd, LB_GETHORIZONTALEXTENT, 0&, ByVal 0&)
    Else
        ScrollWidth = m_ScrollWidth
    End If
End Property
Public Property Let ScrollWidth(ByVal newValue As Long)
    If newValue >= 0 Then
        If m_hWnd <> 0 And (m_ScrollBars = vbHorizontal Or m_ScrollBars = vbBoth) And m_Columns = 0 Then
            SendMessageW m_hWnd, LB_SETHORIZONTALEXTENT, newValue, ByVal 0&
        End If
        m_ScrollWidth = newValue
        If Not blnDesignTime Then Else PropertyChanged "ScrollWidth"
    End If
End Property
Public Property Get SelAnchor() As Long
Attribute SelAnchor.VB_MemberFlags = "400"
    If m_hWnd Then
        SelAnchor = SendMessageW(m_hWnd, LB_GETANCHORINDEX, 0&, ByVal 0&)
    End If
End Property
Public Property Let SelAnchor(ByVal newValue As Long)
    If m_hWnd Then
        SendMessageW m_hWnd, LB_SETANCHORINDEX, newValue, ByVal 0&
    End If
End Property
Public Function SelCount() As Long
    If m_hWnd Then
        SelCount = SendMessageW(m_hWnd, LB_GETSELCOUNT, 0&, ByVal 0&)
    End If
End Function
Public Property Get Selected(ByVal Index As Long) As Boolean
    If m_hWnd Then
        Selected = SendMessageW(m_hWnd, LB_GETSEL, Index, ByVal 0&) <> 0&
    End If
End Property
Public Property Let Selected(ByVal Index As Long, ByVal newValue As Boolean)
    If m_hWnd Then
        SendMessageW m_hWnd, LB_SETSEL, CLng(newValue), ByVal Index
    End If
End Property
Public Sub SetStorage(ByVal Items As Long, ByVal Kilobytes As Long)
    If m_hWnd Then SendMessageW m_hWnd, LB_INITSTORAGE, Items, ByVal Kilobytes * 1024
End Sub
Public Function SetTabStops(ByRef TabStopArray() As Long) As Boolean
    Dim lngTabCount As Long
    If m_hWnd Then
        If Not ((Not TabStopArray) = -1&) Then
            lngTabCount = UBound(TabStopArray) - LBound(TabStopArray) + 1
            SetTabStops = SendMessageW(m_hWnd, LB_SETTABSTOPS, lngTabCount, TabStopArray(LBound(TabStopArray))) <> 0&
        Else
            SetTabStops = SendMessageW(m_hWnd, LB_SETTABSTOPS, 0&, ByVal 0&)
        End If
        On Error Resume Next: Debug.Assert True Xor CBool(0.1): On Error GoTo 0
    End If
End Function
Public Property Get Sort() As Boolean
    Sort = m_Sort
End Property
Public Property Let Sort(ByVal newValue As Boolean)
    If m_Sort <> newValue Then
        m_Sort = newValue
        Private_Init
        If Not blnDesignTime Then Else PropertyChanged "Sort"
    End If
End Property
Public Property Get SortLocale() As UniListLocale
Attribute SortLocale.VB_MemberFlags = "400"
    If m_hWnd Then SortLocale = (SendMessageW(m_hWnd, LB_GETLOCALE, 0&, ByVal 0&) And &HFFFF&)
End Property
Public Property Let SortLocale(ByVal newValue As UniListLocale)
    If m_hWnd Then
        If SendMessageW(m_hWnd, LB_SETLOCALE, &H10000 Or (newValue And &HFFFF&), ByVal 0&) = -1& Then
            If SendMessageW(m_hWnd, LB_SETLOCALE, (newValue And &HFFFF&), ByVal 0&) = -1& Then
                Exit Property
            End If
        End If
        If Not blnDesignTime Then Else PropertyChanged "SortLocale"
    End If
End Property
Public Sub Storage()
    Dim lngBytes As Long
    If m_hWnd Then
        lngBytes = (1048576 * CDbl(m_StorageMB)) And &H7FFFFFFF
        SendMessageW m_hWnd, LB_INITSTORAGE, m_StorageItems, ByVal lngBytes
    End If
End Sub
Public Property Get StorageItems() As Long
    StorageItems = m_StorageItems
End Property
Public Property Let StorageItems(ByVal newValue As Long)
    If newValue >= 100 Then
        m_StorageItems = newValue
        If blnDesignTime Then Else PropertyChanged "StorageItems"
    End If
End Property
Public Property Get StorageMB() As Single
    StorageMB = m_StorageMB
End Property
Public Property Let StorageMB(ByVal newValue As Single)
    If newValue > 512 Then newValue = 512
    If newValue > 0 Then
        m_StorageMB = newValue
        If blnDesignTime Then Else PropertyChanged "StorageMB"
    End If
End Property
Public Property Get Style() As ListBoxConstants
    Style = m_Style
End Property
Public Property Let Style(ByVal NewStyle As ListBoxConstants)
    m_Style = NewStyle
End Property
Public Property Get Text() As String
Attribute Text.VB_MemberFlags = "400"
    Dim bytTemp() As Byte, lngLen As Long, lngListIndex As Long
    If m_hWnd Then
        lngListIndex = SendMessageW(m_hWnd, LB_GETCARETINDEX, 0&, ByVal 0&)
        If lngListIndex >= 0 Then
            lngLen = SendMessageW(m_hWnd, LB_GETTEXTLEN, lngListIndex, ByVal 0&)
            If lngLen > 0 Then
                ReDim bytTemp(lngLen * 2 - 1)
                If SendMessageW(m_hWnd, LB_GETTEXT, lngListIndex, ByVal VarPtr(bytTemp(0))) Then Text = CStr(bytTemp)
            End If
        End If
    End If
End Property
Public Property Let Text(ByRef newValue As String)
    Dim lngListIndex As Long
    If m_hWnd Then
        lngListIndex = SendMessageW(m_hWnd, LB_GETCARETINDEX, 0&, ByVal 0&)
        If lngListIndex >= 0 Then
            If LenB(newValue) Then
                SendMessageW m_hWnd, LB_SETTEXT, lngListIndex, ByVal StrPtr(newValue)
            Else
                SendMessageW m_hWnd, LB_SETTEXT, lngListIndex, ByVal 0&
            End If
        End If
    End If
End Property
Friend Function TranslateAccel(pMsg As Msg) As Boolean
    Dim pOleObject As IOleObject
    Dim pOleControlSite As IOleControlSite
    Dim lngShiftState As Long, bytKeyState(255) As Byte, bytOldState As Byte, intChar As Integer
    
    If GetFocus <> m_hWnd Then SetFocus m_hWnd  'Private_WndEventFocus:
    
    Select Case pMsg.message
        Case WM_KEYDOWN, WM_KEYUP
            Select Case pMsg.wParam
                Case vbKeyTab
                    If m_CaptureTab Then
                        lngShiftState = Private_GetShiftState
                        ' Ctrl + Tab & Shift + Tab move focus out of control
                        If (lngShiftState And vbCtrlMask) = vbCtrlMask Or (lngShiftState And vbShiftMask) = vbShiftMask Then
                            Set pOleObject = Me
                            Set pOleControlSite = pOleObject.GetClientSite
                            If Not pOleControlSite Is Nothing Then
                                pOleControlSite.TranslateAccelerator VarPtr(pMsg), lngShiftState And vbShiftMask
                            End If
                        Else
                            ' hack Ctrl key so that the tab works
                            GetKeyboardState bytKeyState(0)
                            bytOldState = bytKeyState(vbKeyControl)
                            bytKeyState(vbKeyControl) = bytKeyState(vbKeyTab)
                            SetKeyboardState bytKeyState(0)
                            ' send tab
                            SendMessageW m_hWnd, pMsg.message, pMsg.wParam, pMsg.lParam
                            ' restore original Ctrl key state
                            bytKeyState(vbKeyControl) = bytOldState
                            SetKeyboardState bytKeyState(0)
                        End If
                        ' Ignore the message
                        TranslateAccel = True
                    End If
                Case vbKeyReturn
                    If m_CaptureEnter Then
                        SendMessageW m_hWnd, pMsg.message, pMsg.wParam, pMsg.lParam
                        If pMsg.message = WM_KEYDOWN Then RaiseEvent KeyPress(vbKeyReturn)
                        ' Ignore the message
                        TranslateAccel = True
                    End If
                    'End If
                Case vbKeyEscape
                    If m_CaptureEsc Then
                        SendMessageW m_hWnd, pMsg.message, pMsg.wParam, pMsg.lParam
                        If pMsg.message = WM_KEYDOWN Then RaiseEvent KeyPress(vbKeyEscape)
                        ' Ignore the message
                        TranslateAccel = True
                    End If
                Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, _
                     vbKeyEnd, vbKeyPageDown, vbKeyPageUp
                    ' Navigation keys filter
                    SendMessageW m_hWnd, pMsg.message, pMsg.wParam, pMsg.lParam
                    TranslateAccel = True
                Case Else
                    ' if locked prevent every key
                    TranslateAccel = True
            End Select
        Case WM_CHAR
            ' If the control is read-only
            ' eat every char
            If Not m_Enabled Then TranslateAccel = True
    End Select
End Function
Public Property Get TopIndex() As Long
Attribute TopIndex.VB_MemberFlags = "400"
    If m_hWnd Then
        TopIndex = SendMessageW(m_hWnd, LB_GETTOPINDEX, 0&, ByVal 0&)
    End If
End Property
Public Property Let TopIndex(ByVal newValue As Long)
    If m_hWnd Then
        SendMessageW m_hWnd, LB_SETTOPINDEX, newValue, ByVal 0&
        If Not blnDesignTime Then Else PropertyChanged "TopIndex"
    End If
End Property
Public Property Get UseEvents() As Boolean
    UseEvents = m_UseEvents
End Property
Public Property Let UseEvents(ByVal newValue As Boolean)
    m_UseEvents = newValue
End Property
Public Property Get UseTabStops() As Boolean
    UseTabStops = m_UseTabStops
End Property
Public Property Let UseTabStops(ByVal newValue As Boolean)
    If m_UseTabStops <> newValue Then
        m_UseTabStops = newValue
        Private_Init
        If Not blnDesignTime Then Else PropertyChanged "UseTabStops"
    End If
End Property
Private Sub m_Font_FontChanged(ByVal PropertyName As String)
    Dim objFont As IFont
    If m_hWnd Then
        If m_Font Is Nothing Then Exit Sub
        Set objFont = m_Font
        SendMessageW m_hWnd, WM_SETFONT, objFont.hFont, ByVal 0&
        MoveWindow m_hWnd, m_BorderStyle, m_BorderStyle, m_RC.Right - m_BorderStyle * 2, m_RC.Bottom - m_BorderStyle * 2, -1&
        RaiseEvent FontChanged
    End If
End Sub
Private Sub UserControl_AmbientChanged(PropertyName As String)
    ' see if container scalemode has changed
    If LenB(PropertyName) = 20 Then If PropertyName = "ScaleUnits" Then m_ContainerScaleMode = Private_GetContainerScaleMode
End Sub
Private Sub UserControl_EnterFocus()
    '
End Sub
Private Sub UserControl_ExitFocus()
    '
End Sub
Private Sub UserControl_GotFocus()
    If blnDesignTime Then m_Focus = True
    Private_SetIPAO
    If m_hWnd Then If GetFocus <> m_hWnd Then SetFocus m_hWnd
End Sub
Private Sub UserControl_Initialize()
    If Not App.LogMode = 0 Then
        ' this will fix a problem with some versions of comctl32.dll when using XP Themes
        ' http://www.vbaccelerator.com/home/vb/Code/Libraries/XP_Visual_Styles/Preventing_Crashes_at_Shutdown/article.asp
        m_Shell32 = GetModuleHandleA("shell32.dll")
        If m_Shell32 = 0 Then m_Shell32 = LoadLibraryA("shell32.dll"): m_FreeShell32 = True
    End If
    ' see if mouse tracking is supported (WM_MOUSELEAVE)
    m_TrackUser32 = Private_IsFunctionSupported("TrackMouseEvent", "user32.dll")
    If Not m_TrackUser32 Then m_TrackComCtl = Private_IsFunctionSupported("_TrackMouseEvent", "comctl32.dll")
End Sub
Private Sub UserControl_InitProperties()
    ' design time?
    blnDesignTime = (Not Ambient.UserMode)
    If Not blnDesignTime Then UniList_Init m_IPAO, Me
    ' container scalemode
    m_ContainerScaleMode = Private_GetContainerScaleMode
    ' default property values
    m_Appearance = [Classic 3D]
    If UserControl.BackColor < 0 Then m_BackClr = GetSysColor(UserControl.BackColor And &HFF&) Else m_BackClr = UserControl.BackColor
    m_BackClrBrush = CreateSolidBrush(m_BackClr)
    m_BorderStyle = [3D]
    m_Enabled = True
    Set m_Font = Ambient.Font
    m_ForeColor = vbWindowText
    If m_ForeColor < 0 Then m_ForeClr = GetSysColor(m_ForeColor And &HFF&) Else m_ForeClr = m_ForeColor
    m_IntegralHeight = True
    m_ScrollBars = vbVertical
    m_StorageItems = 500
    m_StorageMB = 1
    m_UseEvents = True
    m_UseTabStops = True
    ' subclass
    If ssc_Subclass(UserControl.hWnd, , 1, , Ambient.UserMode, True) Then
        ssc_AddMsg UserControl.hWnd, MSG_BEFORE, WM_COMMAND, WM_DESTROY, WM_ERASEBKGND, WM_CTLCOLOREDIT, WM_CTLCOLORSTATIC, WM_SETFOCUS
        ssc_AddMsg UserControl.hWnd, MSG_AFTER, WM_COMMAND, WM_PAINT
        If m_TrackUser32 Or m_TrackComCtl Then ssc_AddMsg UserControl.hWnd, MSG_AFTER, WM_MOUSELEAVE
        Debug.Print Ambient.DisplayName & ": Started subclassing! " & Hex$(UserControl.hWnd)
    Else
        Debug.Print Ambient.DisplayName & ": Failed to subclass! " & Hex$(UserControl.hWnd)
    End If
    ' initial show
    Private_Init
End Sub
Private Sub UserControl_LostFocus()
    If blnDesignTime Then m_Focus = False
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ' design time?
    blnDesignTime = (Not Ambient.UserMode)
    If Not blnDesignTime Then UniList_Init m_IPAO, Me
    ' container scalemode
    m_ContainerScaleMode = Private_GetContainerScaleMode
    ' load property values
    UserControl.BackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
    If UserControl.BackColor < 0 Then m_BackClr = GetSysColor(UserControl.BackColor And &HFF&) Else m_BackClr = UserControl.BackColor
    m_BackClrBrush = CreateSolidBrush(m_BackClr)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", [3D])
    m_CaptureEnter = PropBag.ReadProperty("CaptureEnter", True)
    m_CaptureEsc = PropBag.ReadProperty("CaptureEsc", False)
    m_CaptureTab = PropBag.ReadProperty("CaptureTab", False)
    m_Columns = PropBag.ReadProperty("Columns", 0)
    m_DisableSelect = PropBag.ReadProperty("DisableSelect", False)
    m_Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.Enabled = m_Enabled
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", vbWindowText)
    If m_ForeColor < 0 Then m_ForeClr = GetSysColor(m_ForeColor And &HFF&) Else m_ForeClr = m_ForeColor
    m_IntegralHeight = PropBag.ReadProperty("IntegralHeight", True)
    Set UserControl.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", vbDefault)
    m_MultiSelect = PropBag.ReadProperty("MultiSelect", vbMultiSelectNone)
    m_RightToLeft = PropBag.ReadProperty("RightToLeft", False)
    m_ScrollBars = PropBag.ReadProperty("ScrollBars", vbSBNone)
    m_ScrollBarVisibility = PropBag.ReadProperty("ScrollBarVisibility", [Allow No ScrollBars])
    m_ScrollWidth = PropBag.ReadProperty("ScrollWidth", 0)
    m_Sort = PropBag.ReadProperty("Sort", False)
    m_StorageItems = PropBag.ReadProperty("StorageItems", 500)
    m_StorageMB = PropBag.ReadProperty("StorageMB", 1)
    m_Style = PropBag.ReadProperty("Style", vbListBoxStandard)
    m_UseEvents = PropBag.ReadProperty("UseEvents", True)
    m_UseTabStops = PropBag.ReadProperty("UseTabStops", True)
    ' subclass
    If ssc_Subclass(UserControl.hWnd, , 1, , Ambient.UserMode, True) Then
        ssc_AddMsg UserControl.hWnd, MSG_BEFORE, WM_COMMAND, WM_DESTROY, WM_ERASEBKGND, WM_CTLCOLOREDIT, WM_CTLCOLORSTATIC, WM_SETFOCUS
        ssc_AddMsg UserControl.hWnd, MSG_AFTER, WM_COMMAND, WM_PAINT
        If m_TrackUser32 Or m_TrackComCtl Then ssc_AddMsg UserControl.hWnd, MSG_AFTER, WM_MOUSELEAVE
        Debug.Print Ambient.DisplayName & ": Started subclassing! " & Hex$(UserControl.hWnd)
    Else
        Debug.Print Ambient.DisplayName & ": Failed to subclass! " & Hex$(UserControl.hWnd)
    End If
    ' initial show
    Private_Init
End Sub
Private Sub UserControl_Resize()
    Static blnResize As Boolean
    If blnResize Then Exit Sub
    blnResize = True
    ' update rectangle
    m_RC.Right = UserControl.ScaleWidth
    m_RC.Bottom = UserControl.ScaleHeight
    ' exit if missing
    If m_hWnd Then
        If m_Appearance = [Classic 3D] Then
            ' borderstyle value tells us how deep we go...
            MoveWindow m_hWnd, m_BorderStyle, m_BorderStyle, m_RC.Right - m_BorderStyle * 2, m_RC.Bottom - m_BorderStyle * 2, -1&
            If m_IntegralHeight Then
                GetWindowRect m_hWnd, m_RC
                m_RC.Right = m_RC.Right - m_RC.Left + m_BorderStyle * 2
                m_RC.Bottom = m_RC.Bottom - m_RC.Top + m_BorderStyle * 2
                m_RC.Left = 0
                m_RC.Top = 0
                UserControl.Height = UserControl.ScaleY(m_RC.Bottom, vbPixels, vbTwips)
            End If
        Else
            MoveWindow m_hWnd, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, -1&
            If m_IntegralHeight Then
                GetWindowRect m_hWnd, m_RC
                m_RC.Right = m_RC.Right - m_RC.Left
                m_RC.Bottom = m_RC.Bottom - m_RC.Top
                m_RC.Left = 0
                m_RC.Top = 0
                UserControl.Height = UserControl.ScaleY(m_RC.Bottom, vbPixels, vbTwips)
            End If
        End If
        If (m_Columns > 0) And (m_hWnd <> 0) Then
            SendMessageW m_hWnd, LB_SETCOLUMNWIDTH, (m_RC.Right - m_BorderStyle) \ m_Columns, ByVal 0&
        End If
    End If
    blnResize = False
End Sub

Private Sub UserControl_Terminate()
    UniList_Terminate m_IPAO
    ' remove backcolor
    If m_BackClrBrush Then DeleteObject m_BackClrBrush
    ' remove listbox
    If m_hWnd Then
        ssc_UnSubclass m_hWnd
        DestroyWindow m_hWnd
        m_hWnd = 0
    End If
    ' unload shell32 if it was loaded by this control
    If m_FreeShell32 Then FreeLibrary m_Shell32
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    ' save settings
    PropBag.WriteProperty "BackColor", UserControl.BackColor
    PropBag.WriteProperty "BorderStyle", m_BorderStyle
    PropBag.WriteProperty "CaptureEnter", m_CaptureEnter
    PropBag.WriteProperty "CaptureEsc", m_CaptureEsc
    PropBag.WriteProperty "CaptureTab", m_CaptureTab
    PropBag.WriteProperty "Columns", m_Columns
    PropBag.WriteProperty "DisableSelect", m_DisableSelect
    PropBag.WriteProperty "Enabled", m_Enabled
    PropBag.WriteProperty "Font", m_Font
    PropBag.WriteProperty "ForeColor", m_ForeColor
    PropBag.WriteProperty "IntegralHeight", m_IntegralHeight
    PropBag.WriteProperty "MouseIcon", UserControl.MouseIcon
    PropBag.WriteProperty "MousePointer", UserControl.MousePointer
    PropBag.WriteProperty "MultiSelect", m_MultiSelect
    PropBag.WriteProperty "RightToLeft", m_RightToLeft
    PropBag.WriteProperty "ScrollBars", m_ScrollBars
    PropBag.WriteProperty "ScrollBarVisibility", m_ScrollBarVisibility
    PropBag.WriteProperty "ScrollWidth", m_ScrollWidth
    PropBag.WriteProperty "Sort", m_Sort
    PropBag.WriteProperty "StorageItems", m_StorageItems
    PropBag.WriteProperty "StorageMB", m_StorageMB
    PropBag.WriteProperty "Style", m_Style
    PropBag.WriteProperty "UseEvents", m_UseEvents
    PropBag.WriteProperty "UseTabStops", m_UseTabStops
End Sub

'-SelfSub code------------------------------------------------------------------------------------
'-The following routines are exclusively for the ssc_Subclass routines----------------------------
Private Function ssc_Subclass(ByVal lng_hWnd As Long, _
                    Optional ByVal lParamUser As Long = 0, _
                    Optional ByVal nOrdinal As Long = 1, _
                    Optional ByVal oCallback As Object = Nothing, _
                    Optional ByVal bIdeSafety As Boolean = True, _
                    Optional ByRef bUnicode As Boolean = False, _
                    Optional ByVal bIsAPIwindow As Boolean = False) As Boolean 'Subclass the specified window handle

    '*************************************************************************************************
    '* lng_hWnd   - Handle of the window to subclass
    '* lParamUser - Optional, user-defined callback parameter
    '* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
    '* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety - Optional, enable/disable IDE safety measures. There is not reason to set this to False
    '* bUnicode - Optional, if True, Unicode API calls should be made to the window vs ANSI calls
    '*            Parameter is byRef and its return value should be checked to know if ANSI to be used or not
    '* bIsAPIwindow - Optional, if True DestroyWindow will be called if IDE ENDs
    '*****************************************************************************************
    '** Subclass.asm - subclassing thunk
    '**
    '** Paul_Caton@hotmail.com
    '** Copyright free, use and abuse as you see fit.
    '**
    '** v2.0 Re-write by LaVolpe, based mostly on Paul Caton's original thunks....... 20070720
    '** .... Reorganized & provided following additional logic
    '** ....... Unsubclassing only occurs after thunk is no longer recursed
    '** ....... Flag used to bypass callbacks until unsubclassing can occur
    '** ....... Timer used as delay mechanism to free thunk memory afer unsubclassing occurs
    '** .............. Prevents crash when one window subclassed multiple times
    '** .............. More END safe, even if END occurs within the subclass procedure
    '** ....... Added ability to destroy API windows when IDE terminates
    '** ....... Added auto-unsubclass when WM_NCDESTROY received
    '*****************************************************************************************
    ' Subclassing procedure must be declared identical to the one at the end of this class (Sample at Ordinal #1)

    Dim z_Sc(0 To IDX_UNICODE) As Long                 'Thunk machine-code initialised here
    
    Const SUB_NAME      As String = "ssc_Subclass"     'This routine's name
    Const CODE_LEN      As Long = 4 * IDX_UNICODE + 4  'Thunk length in bytes
    Const PAGE_RWX      As Long = &H40&                'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&              'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&              'Release allocated memory flag
    Const GWL_WNDPROC   As Long = -4                   'SetWindowsLong WndProc index
    Const WNDPROC_OFF   As Long = &H60                 'Thunk offset to the WndProc execution address
    Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1)) 'Bytes to allocate per thunk, data + code + msg tables
    
  ' This is the complete listing of thunk offset values and what they point/relate to.
  ' Those rem'd out are used elsewhere or are initialized in Declarations section
  
  'Const IDX_RECURSION  As Long = 0     'Thunk data index of callback recursion count
  'Const IDX_SHUTDOWN   As Long = 1     'Thunk data index of the termination flag
  'Const IDX_INDEX      As Long = 2     'Thunk data index of the subclassed hWnd
   Const IDX_EBMODE     As Long = 3     'Thunk data index of the EbMode function address
   Const IDX_CWP        As Long = 4     'Thunk data index of the CallWindowProc function address
   Const IDX_SWL        As Long = 5     'Thunk data index of the SetWindowsLong function address
   Const IDX_FREE       As Long = 6     'Thunk data index of the VirtualFree function address
   Const IDX_BADPTR     As Long = 7     'Thunk data index of the IsBadCodePtr function address
   Const IDX_OWNER      As Long = 8     'Thunk data index of the Owner object's vTable address
  'Const IDX_PREVPROC   As Long = 9     'Thunk data index of the original WndProc
   Const IDX_CALLBACK   As Long = 10    'Thunk data index of the callback method address
  'Const IDX_BTABLE     As Long = 11    'Thunk data index of the Before table
  'Const IDX_ATABLE     As Long = 12    'Thunk data index of the After table
  'Const IDX_PARM_USER  As Long = 13    'Thunk data index of the User-defined callback parameter data index
   Const IDX_DW         As Long = 14    'Thunk data index of the DestroyWinodw function address
   Const IDX_ST         As Long = 15    'Thunk data index of the SetTimer function address
   Const IDX_KT         As Long = 16    'Thunk data index of the KillTimer function address
   Const IDX_EBX_TMR    As Long = 20    'Thunk code patch index of the thunk data for the delay timer
   Const IDX_EBX        As Long = 26    'Thunk code patch index of the thunk data
  'Const IDX_UNICODE    As Long = xx    'Must be UBound(subclass thunkdata)+1; index for unicode support
    
    Dim z_ScMem       As Long           'Thunk base address
    Dim nAddr         As Long
    Dim nID           As Long
    Dim nMyID         As Long
    Dim bIDE          As Boolean

    If IsWindow(lng_hWnd) = 0 Then      'Ensure the window handle is valid
        zError SUB_NAME, "Invalid window handle"
        Exit Function
    End If
    
    nMyID = GetCurrentProcessId                         'Get this process's ID
    GetWindowThreadProcessId lng_hWnd, nID              'Get the process ID associated with the window handle
    If nID <> nMyID Then                                'Ensure that the window handle doesn't belong to another process
        zError SUB_NAME, "Window handle belongs to another process"
        Exit Function
    End If
    
    If oCallback Is Nothing Then Set oCallback = Me     'If the user hasn't specified the callback owner
    
    nAddr = zAddressOf(oCallback, nOrdinal)             'Get the address of the specified ordinal method
    If nAddr = 0 Then                                   'Ensure that we've found the ordinal method
        zError SUB_NAME, "Callback method not found"
        Exit Function
    End If
        
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
    
    If z_ScMem <> 0 Then                                  'Ensure the allocation succeeded
    
      If z_scFunk Is Nothing Then Set z_scFunk = New Collection 'If this is the first time through, do the one-time initialization
      On Error GoTo CatchDoubleSub                              'Catch double subclassing
        z_scFunk.Add z_ScMem, "h" & lng_hWnd                    'Add the hWnd/thunk-address to the collection
      On Error GoTo 0
      
   'z_Sc (0) thru z_Sc(17) are used as storage for the thunks & IDX_ constants above relate to these thunk positions which are filled in below
    z_Sc(18) = &HD231C031: z_Sc(19) = &HBBE58960: z_Sc(21) = &H21E8F631: z_Sc(22) = &HE9000001: z_Sc(23) = &H12C&: z_Sc(24) = &HD231C031: z_Sc(25) = &HBBE58960: z_Sc(27) = &H3FFF631: z_Sc(28) = &H75047339: z_Sc(29) = &H2873FF23: z_Sc(30) = &H751C53FF: z_Sc(31) = &HC433913: z_Sc(32) = &H53FF2274: z_Sc(33) = &H13D0C: z_Sc(34) = &H18740000: z_Sc(35) = &H875C085: z_Sc(36) = &H820443C7: z_Sc(37) = &H90000000: z_Sc(38) = &H87E8&: z_Sc(39) = &H22E900: z_Sc(40) = &H90900000: z_Sc(41) = &H2C7B8B4A: z_Sc(42) = &HE81C7589: z_Sc(43) = &H90&: z_Sc(44) = &H75147539: z_Sc(45) = &H6AE80F: z_Sc(46) = &HD2310000: z_Sc(47) = &HE8307B8B: z_Sc(48) = &H7C&: z_Sc(49) = &H7D810BFF: z_Sc(50) = &H8228&: z_Sc(51) = &HC7097500: z_Sc(52) = &H80000443: z_Sc(53) = &H90900000: z_Sc(54) = &H44753339: z_Sc(55) = &H74047339: z_Sc(56) = &H2473FF3F: z_Sc(57) = &HFFFFFC68
    z_Sc(58) = &H2475FFFF: z_Sc(59) = &H811453FF: z_Sc(60) = &H82047B: z_Sc(61) = &HC750000: z_Sc(62) = &H74387339: z_Sc(63) = &H2475FF07: z_Sc(64) = &H903853FF: z_Sc(65) = &H81445B89: z_Sc(66) = &H484443: z_Sc(67) = &H73FF0000: z_Sc(68) = &H646844: z_Sc(69) = &H56560000: z_Sc(70) = &H893C53FF: z_Sc(71) = &H90904443: z_Sc(72) = &H10C261: z_Sc(73) = &H53E8&: z_Sc(74) = &H3075FF00: z_Sc(75) = &HFF2C75FF: z_Sc(76) = &H75FF2875: z_Sc(77) = &H2473FF24: z_Sc(78) = &H891053FF: z_Sc(79) = &H90C31C45: z_Sc(80) = &H34E30F8B: z_Sc(81) = &H1078C985: z_Sc(82) = &H4C781: z_Sc(83) = &H458B0000: z_Sc(84) = &H75AFF228: z_Sc(85) = &H90909023: z_Sc(86) = &H8D144D8D: z_Sc(87) = &H8D503443: z_Sc(88) = &H75FF1C45: z_Sc(89) = &H2C75FF30: z_Sc(90) = &HFF2875FF: z_Sc(91) = &H51502475: z_Sc(92) = &H2073FF52: z_Sc(93) = &H902853FF: z_Sc(94) = &H909090C3: z_Sc(95) = &H74447339: z_Sc(96) = &H4473FFF7
    z_Sc(97) = &H4053FF56: z_Sc(98) = &HC3447389: z_Sc(99) = &H89285D89: z_Sc(100) = &H45C72C75: z_Sc(101) = &H800030: z_Sc(102) = &H20458B00: z_Sc(103) = &H89145D89: z_Sc(104) = &H81612445: z_Sc(105) = &H4C4&: z_Sc(106) = &H1862FF00

    ' cache callback related pointers & offsets
      z_Sc(IDX_EBX) = z_ScMem                                                 'Patch the thunk data address
      z_Sc(IDX_EBX_TMR) = z_ScMem                                             'Patch the thunk data address
      z_Sc(IDX_INDEX) = lng_hWnd                                              'Store the window handle in the thunk data
      z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                   'Store the address of the before table in the thunk data
      z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
      z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback owner's object address in the thunk data
      z_Sc(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data
      z_Sc(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data
      
      ' validate unicode request & cache unicode usage
      If bUnicode Then bUnicode = (IsWindowUnicode(lng_hWnd) <> 0&)
      z_Sc(IDX_UNICODE) = bUnicode                                            'Store whether the window is using unicode calls or not
      
      ' get function pointers for the thunk
      If bIdeSafety = True Then                                               'If the user wants IDE protection
          Debug.Assert zInIDE(bIDE)
          If bIDE = True Then z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode", bUnicode) 'Store the EbMode function address in the thunk data
                                                        '^^ vb5 users, change vba6 to vba5
      End If
      If bIsAPIwindow Then                                                    'If user wants DestroyWindow sent should IDE end
          z_Sc(IDX_DW) = zFnAddr("user32", "DestroyWindow", bUnicode)
      End If
      z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree", bUnicode)           'Store the VirtualFree function address in the thunk data
      z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", bUnicode)        'Store the IsBadCodePtr function address in the thunk data
      z_Sc(IDX_ST) = zFnAddr("user32", "SetTimer", bUnicode)                  'Store the SetTimer function address in the thunk data
      z_Sc(IDX_KT) = zFnAddr("user32", "KillTimer", bUnicode)                 'Store the KillTimer function address in the thunk data
      
      If bUnicode Then
          z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcW", bUnicode)      'Store CallWindowProc function address in the thunk data
          z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongW", bUnicode)       'Store the SetWindowLong function address in the thunk data
          RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                    'Copy the thunk code/data to the allocated memory
          z_Sc(IDX_PREVPROC) = SetWindowLongW(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Set the new WndProc, return the address of the original WndProc
      Else
          z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA", bUnicode)      'Store CallWindowProc function address in the thunk data
          z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA", bUnicode)       'Store the SetWindowLong function address in the thunk data
          RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                    'Copy the thunk code/data to the allocated memory
          z_Sc(IDX_PREVPROC) = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Set the new WndProc, return the address of the original WndProc
      End If
      If z_Sc(IDX_PREVPROC) = 0 Then                                          'Ensure the new WndProc was set correctly
          zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
          GoTo ReleaseMemory
      End If
      'Store the original WndProc address in the thunk data
      RtlMoveMemory z_ScMem + IDX_PREVPROC * 4, VarPtr(z_Sc(IDX_PREVPROC)), 4&
      ssc_Subclass = True                                                     'Indicate success
      
    Else
        zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
        
    End If

 Exit Function                                                                'Exit ssc_Subclass
    
CatchDoubleSub:
 zError SUB_NAME, "Window handle is already subclassed"
      
ReleaseMemory:
      VirtualFree z_ScMem, 0, MEM_RELEASE                                     'ssc_Subclass has failed after memory allocation, so release the memory
      
End Function

'Terminate all subclassing
Private Sub ssc_Terminate()
    ' can be made public, can be removed & zTerminateThunks can be called instead
    zTerminateThunks SubclassThunk
End Sub

'UnSubclass the specified window handle
Private Sub ssc_UnSubclass(ByVal lng_hWnd As Long)
    ' can be made public, can be removed & zUnthunk can be called instead
    zUnThunk lng_hWnd, SubclassThunk
End Sub

'Add the message value to the window handle's specified callback table
Private Sub ssc_AddMsg(ByVal lng_hWnd As Long, ByVal When As eMsgWhen, ParamArray Messages() As Variant)
    
    Dim z_ScMem       As Long                                   'Thunk base address
    
    z_ScMem = zMap_VFunction(lng_hWnd, SubclassThunk)           'Ensure that the thunk hasn't already released its memory
    If z_ScMem Then
      Dim M As Long
      For M = LBound(Messages) To UBound(Messages)
        Select Case VarType(Messages(M))                        ' ensure no strings, arrays, doubles, objects, etc are passed
        Case vbByte, vbInteger, vbLong
            If When And MSG_BEFORE Then                         'If the message is to be added to the before original WndProc table...
              If zAddMsg(Messages(M), IDX_BTABLE, z_ScMem) = False Then 'Add the message to the before table
                When = (When And Not MSG_BEFORE)
              End If
            End If
            If When And MSG_AFTER Then                          'If message is to be added to the after original WndProc table...
              If zAddMsg(Messages(M), IDX_ATABLE, z_ScMem) = False Then 'Add the message to the after table
                When = (When And Not MSG_AFTER)
              End If
            End If
        End Select
      Next
    End If
End Sub

'Delete the message value from the window handle's specified callback table
Private Sub ssc_DelMsg(ByVal lng_hWnd As Long, ByVal When As eMsgWhen, ParamArray Messages() As Variant)
    
    Dim z_ScMem       As Long                                                   'Thunk base address
    
    z_ScMem = zMap_VFunction(lng_hWnd, SubclassThunk)                           'Ensure that the thunk hasn't already released its memory
    If z_ScMem Then
      Dim M As Long
      For M = LBound(Messages) To UBound(Messages)                              ' ensure no strings, arrays, doubles, objects, etc are passed
        Select Case VarType(Messages(M))
        Case vbByte, vbInteger, vbLong
            If When And MSG_BEFORE Then                                         'If the message is to be removed from the before original WndProc table...
              zDelMsg Messages(M), IDX_BTABLE, z_ScMem                          'Remove the message to the before table
            End If
            If When And MSG_AFTER Then                                          'If message is to be removed from the after original WndProc table...
              zDelMsg Messages(M), IDX_ATABLE, z_ScMem                          'Remove the message to the after table
            End If
        End Select
      Next
    End If
End Sub

'Call the original WndProc
Private Function ssc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' can be made public, can be removed if you will not use this in your window procedure
    Dim z_ScMem       As Long                           'Thunk base address
    z_ScMem = zMap_VFunction(lng_hWnd, SubclassThunk)
    If z_ScMem Then                                     'Ensure that the thunk hasn't already released its memory
        If zData(IDX_UNICODE, z_ScMem) Then
            ssc_CallOrigWndProc = CallWindowProcW(zData(IDX_PREVPROC, z_ScMem), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        Else
            ssc_CallOrigWndProc = CallWindowProcA(zData(IDX_PREVPROC, z_ScMem), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        End If
    End If
End Function

'Get the subclasser lParamUser callback parameter
Private Function zGet_lParamUser(ByVal hWnd_Hook_ID As Long, ByVal vType As eThunkType) As Long
    ' can be removed if you never will retrieve or replace the user-defined parameter
    If vType <> CallbackThunk Then
        Dim z_ScMem       As Long                                       'Thunk base address
        z_ScMem = zMap_VFunction(hWnd_Hook_ID, vType)
        If z_ScMem Then                                                 'Ensure that the thunk hasn't already released its memory
          zGet_lParamUser = zData(IDX_PARM_USER, z_ScMem)               'Get the lParamUser callback parameter
        End If
    End If
End Function

'Let the subclasser lParamUser callback parameter
Private Sub zSet_lParamUser(ByVal hWnd_Hook_ID As Long, ByVal vType As eThunkType, ByVal newValue As Long)
    ' can be removed if you never will retrieve or replace the user-defined parameter
    If vType <> CallbackThunk Then
        Dim z_ScMem       As Long                                       'Thunk base address
        z_ScMem = zMap_VFunction(hWnd_Hook_ID, vType)
        If z_ScMem Then                                                 'Ensure that the thunk hasn't already released its memory
          zData(IDX_PARM_USER, z_ScMem) = newValue                      'Set the lParamUser callback parameter
        End If
    End If
End Sub

'Add the message to the specified table of the window handle
Private Function zAddMsg(ByVal uMsg As Long, ByVal nTable As Long, ByVal z_ScMem As Long) As Boolean
      Dim nCount As Long                                                        'Table entry count
      Dim nBase  As Long
      Dim I      As Long                                                        'Loop index
    
      zAddMsg = True
      nBase = zData(nTable, z_ScMem)                                            'Map zData() to the specified table
      
      If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
        nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
      Else
        
        nCount = zData(0, nBase)                                                'Get the current table entry count
        For I = 1 To nCount                                                     'Loop through the table entries
          If zData(I, nBase) = 0 Then                                           'If the element is free...
            zData(I, nBase) = uMsg                                              'Use this element
            GoTo Bail                                                           'Bail
          ElseIf zData(I, nBase) = uMsg Then                                    'If the message is already in the table...
            GoTo Bail                                                           'Bail
          End If
        Next I                                                                  'Next message table entry
    
        nCount = I                                                             'On drop through: i = nCount + 1, the new table entry count
        If nCount > MSG_ENTRIES Then                                           'Check for message table overflow
          zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
          zAddMsg = False
          GoTo Bail
        End If
        
        zData(nCount, nBase) = uMsg                                            'Store the message in the appended table entry
      End If
    
      zData(0, nBase) = nCount                                                 'Store the new table entry count
Bail:
End Function

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long, ByVal z_ScMem As Long)
      Dim nCount As Long                                                        'Table entry count
      Dim nBase  As Long
      Dim I      As Long                                                        'Loop index
    
      nBase = zData(nTable, z_ScMem)                                            'Map zData() to the specified table
    
      If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
        zData(0, nBase) = 0                                                     'Zero the table entry count
      Else
        nCount = zData(0, nBase)                                                'Get the table entry count
        
        For I = 1 To nCount                                                     'Loop through the table entries
          If zData(I, nBase) = uMsg Then                                        'If the message is found...
            zData(I, nBase) = 0                                                 'Null the msg value -- also frees the element for re-use
            GoTo Bail                                                           'Bail
          End If
        Next I                                                                  'Next message table entry
        
       ' zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
      End If
Bail:
End Sub

'-The following routines are used for each of the three types of thunks ----------------------------

'Maps zData() to the memory address for the specified thunk type
Private Function zMap_VFunction(vFuncTarget As Long, _
                                vType As eThunkType, _
                                Optional oCallback As Object, _
                                Optional bIgnoreErrors As Boolean) As Long
    
    Dim thunkCol As Collection
    Dim colID As String
    Dim z_ScMem       As Long         'Thunk base address
    
    If vType = CallbackThunk Then
        Set thunkCol = z_cbFunk
        If oCallback Is Nothing Then Set oCallback = Me
        colID = "h" & ObjPtr(oCallback) & "." & vFuncTarget
    ElseIf vType = HookThunk Then
        Set thunkCol = z_hkFunk
        colID = "h" & vFuncTarget
    ElseIf vType = SubclassThunk Then
        Set thunkCol = z_scFunk
        colID = "h" & vFuncTarget
    Else
        zError "zMap_Vfunction", "Invalid thunk type passed"
        Exit Function
    End If
    
    If thunkCol Is Nothing Then
        zError "zMap_VFunction", "Thunk hasn't been initialized"
    Else
        If thunkCol.Count Then
            On Error GoTo Catch
            z_ScMem = thunkCol(colID)               'Get the thunk address
            If IsBadCodePtr(z_ScMem) Then z_ScMem = 0&
            zMap_VFunction = z_ScMem
        End If
    End If
    Exit Function                                               'Exit returning the thunk address
    
Catch:
    ' error ignored when zUnThunk is called, error handled there
    If Not bIgnoreErrors Then zError "zMap_VFunction", "Thunk type for " & vType & " does not exist"
End Function

' sets/retrieves data at the specified offset for the specified memory address
Private Property Get zData(ByVal nIndex As Long, ByVal z_ScMem As Long) As Long
  RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal z_ScMem As Long, ByVal nValue As Long)
  RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4
End Property

'Error handler
Private Sub zError(ByRef sRoutine As String, ByVal sMsg As String)
  ' Note. These two lines can be rem'd out if you so desire. But don't remove the routine
  App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String, ByVal asUnicode As Boolean) As Long
  If asUnicode Then
    zFnAddr = GetProcAddress(GetModuleHandleW(StrPtr(sDLL)), sProc)         'Get the specified procedure address
  Else
    zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                 'Get the specified procedure address
  End If
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
  ' ^^ FYI VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
    ' Note: used both in subclassing and hooking routines
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim I     As Long                                                         'Loop index
  Dim j     As Long                                                         'Loop limit
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, I, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, I, bSub) Then                              'Probe for a Form method
      If Not zProbe(nAddr + &H710, I, bSub) Then                            'Probe for a PropertyPage method
        If Not zProbe(nAddr + &H7A4, I, bSub) Then                          'Probe for a UserControl method
            Exit Function                                                   'Bail...
        End If
      End If
    End If
  End If
  
  I = I + 4                                                                 'Bump to the next entry
  j = I + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While I < j
    RtlMoveMemory VarPtr(nAddr), I, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      RtlMoveMemory VarPtr(zAddressOf), I - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      RtlMoveMemory VarPtr(zAddressOf), I - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    I = I + 4                                                               'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Do                                                             'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

Private Function zInIDE(ByRef bIDE As Boolean) As Boolean
    ' only called in IDE, never called when compiled
    bIDE = True
    zInIDE = bIDE
End Function

Private Sub zUnThunk(ByVal thunkID As Long, ByVal vType As eThunkType, Optional ByVal oCallback As Object)

    ' thunkID, depends on vType:
    '   - Subclassing:  the hWnd of the window subclassed
    '   - Hooking:      the hook type created
    '   - Callbacks:    the ordinal of the callback
    '       ensure KillTimer is already called, if any callback used for SetTimer
    ' oCallback only used when vType is CallbackThunk

    Const IDX_SHUTDOWN  As Long = 1
    Const MEM_RELEASE As Long = &H8000&             'Release allocated memory flag
    
    Dim z_ScMem       As Long                       'Thunk base address
    
    z_ScMem = zMap_VFunction(thunkID, vType, oCallback, True)
    Select Case vType
    Case SubclassThunk
        If z_ScMem Then                                 'Ensure that the thunk hasn't already released its memory
            zData(IDX_SHUTDOWN, z_ScMem) = 1            'Set the shutdown indicator
            zDelMsg ALL_MESSAGES, IDX_BTABLE, z_ScMem   'Delete all before messages
            zDelMsg ALL_MESSAGES, IDX_ATABLE, z_ScMem   'Delete all after messages
        End If
        If thunkID <> 0 Then z_scFunk.Remove "h" & thunkID                   'Remove the specified thunk from the collection
        
    Case HookThunk
        If z_ScMem Then                                 'Ensure that the thunk hasn't already released its memory
            ' if not unhooked, then unhook now
            If zData(IDX_SHUTDOWN, z_ScMem) = 0 Then UnHookWindowsHookEx zData(IDX_PREVPROC, z_ScMem)
            If zData(0, z_ScMem) = 0 Then               ' not recursing then
                VirtualFree z_ScMem, 0, MEM_RELEASE     'Release allocated memory
                z_hkFunk.Remove "h" & thunkID           'Remove the specified thunk from the collection
            Else
                zData(IDX_SHUTDOWN, z_ScMem) = 1        ' Set the shutdown indicator
                zData(IDX_ATABLE, z_ScMem) = 0          ' want no more After messages
                zData(IDX_BTABLE, z_ScMem) = 0          ' want no more Before messages
                ' when zTerminate is called this thunk's memory will be released
            End If
        Else
            z_hkFunk.Remove "h" & thunkID       'Remove the specified thunk from the collection
        End If
    Case CallbackThunk
        If z_ScMem Then                         'Ensure that the thunk hasn't already released its memory
            VirtualFree z_ScMem, 0, MEM_RELEASE 'Release allocated memory
        End If
        z_cbFunk.Remove "h" & ObjPtr(oCallback) & "." & thunkID           'Remove the specified thunk from the collection
    End Select

End Sub

Private Sub zTerminateThunks(ByVal vType As eThunkType)

    ' Terminates all thunks of a specific type
    ' Any subclassing, hooking, recurring callbacks should have already been canceled

    Dim I As Long
    Dim oCallback As Object
    Dim thunkCol As Collection
    Dim z_ScMem       As Long                           'Thunk base address
    Const INDX_OWNER As Long = 0
    
    Select Case vType
    Case SubclassThunk
        Set thunkCol = z_scFunk
    Case HookThunk
        Set thunkCol = z_hkFunk
    Case CallbackThunk
        Set thunkCol = z_cbFunk
    Case Else
        Exit Sub
    End Select
    
    If Not (thunkCol Is Nothing) Then                 'Ensure that hooking has been started
      With thunkCol
        For I = .Count To 1 Step -1                   'Loop through the collection of hook types in reverse order
          z_ScMem = .Item(I)                          'Get the thunk address
          If IsBadCodePtr(z_ScMem) = 0 Then           'Ensure that the thunk hasn't already released its memory
            Select Case vType
                Case SubclassThunk
                    zUnThunk zData(IDX_INDEX, z_ScMem), SubclassThunk    'Unsubclass
                Case HookThunk
                    zUnThunk zData(IDX_INDEX, z_ScMem), HookThunk        'Unhook
                Case CallbackThunk
                    ' zUnThunk expects object not pointer, convert pointer to object
                    RtlMoveMemory VarPtr(oCallback), VarPtr(zData(INDX_OWNER, z_ScMem)), 4&
                    zUnThunk zData(IDX_CALLBACKORDINAL, z_ScMem), CallbackThunk, oCallback ' release callback
                    ' remove the object pointer reference
                    RtlMoveMemory VarPtr(oCallback), VarPtr(INDX_OWNER), 4&
            End Select
          End If
        Next I                                        'Next member of the collection
      End With
      Set thunkCol = Nothing                         'Destroy the hook/thunk-address collection
    End If


End Sub

' WNDPROC RELATED PROCEDURES

Private Function Private_WndEventKeyboard(MF As MSGFILTER) As Long
    Dim intShift As Integer
    Dim intChar As Integer
    Select Case MF.Msg
        Case WM_KEYDOWN
            ' set the intShift parameter
            intShift = Private_GetShiftMask
            ' set the KeyCode parameter
            intChar = MF.wParam And &HFFFF&
            ' raise the event
            RaiseEvent KeyDown(intChar, intShift)
        Case WM_KEYUP
            intShift = Private_GetShiftMask
            intChar = MF.wParam And &HFFFF&
            RaiseEvent KeyUp(intChar, intShift)
        Case WM_CHAR
            intChar = MF.wParam And &HFFFF&
            RaiseEvent KeyPress(intChar)
    End Select
    Private_WndEventKeyboard = Abs(intChar = 0)
End Function

Private Sub Private_WndProcList(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lParamUser As Long)
    Static KeyShift As Integer
    Dim tme As TRACKMOUSEEVENT_STRUCT
    Dim XY(1) As Integer, Shift As UniListShift, Button As UniListMouseButton
    Dim NH As NMHDR, MF As MSGFILTER
    
    If blnDesignTime And Not m_Focus Then Exit Sub
    
    If bBefore Then
        If uMsg = WM_MOUSEACTIVATE Then
            If GetFocus <> m_hWnd Then SetFocus UserControl.hWnd
        ElseIf m_UseEvents Then
            Select Case uMsg
                Case WM_KEYDOWN, WM_CHAR
                    MF.Msg = uMsg
                    MF.lParam = lParam
                    MF.wParam = wParam
                    bHandled = Private_WndEventKeyboard(MF) <> 0
                Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN
                    ' this horribility is one reason to not use events...
                    CopyMemory XY(0), lParam, 4&
                    ' see key states
                    Shift = Private_GetShiftMask
                    ' button states
                    If (wParam And MK_LBUTTON) = MK_LBUTTON Then Button = [Left Button]
                    If (wParam And MK_MBUTTON) = MK_MBUTTON Then Button = Button Or [Middle Button]
                    If (wParam And MK_RBUTTON) = MK_RBUTTON Then Button = Button Or [Right Button]
                    ' raise the event
                    RaiseEvent MouseDown(Button, Shift, ScaleX(XY(0), vbPixels, m_ContainerScaleMode), ScaleY(XY(1), vbPixels, m_ContainerScaleMode))
            End Select
        End If
    ' BEFORE
    Else
    ' AFTER
        Select Case uMsg
            Case WM_VSCROLL
                RaiseEvent Scroll([Scroll Vertical])
            Case WM_HSCROLL
                RaiseEvent Scroll([Scroll Horizontal])
            Case WM_SETFOCUS
                Private_SetIPAO
                SetFocus m_hWnd
            Case WM_KEYUP
                MF.Msg = WM_KEYUP
                MF.lParam = lParam
                MF.wParam = wParam
                bHandled = Private_WndEventKeyboard(MF) <> 1
            Case WM_MOUSEMOVE
                ' see if entering into the control
                If Not m_MouseOver Then
                    ' initialize TrackMouseEvent structure
                    tme.cbSize = Len(tme)
                    tme.dwFlags = TME_LEAVE
                    tme.hwndTrack = lng_hWnd
                    ' see which tracking API is available, if any
                    If m_TrackUser32 Then
                        TrackMouseEventUser32 tme
                    ElseIf m_TrackComCtl Then
                        TrackMouseEventComCtl tme
                    End If
                    ' set mouseover
                    m_MouseOver = True
                    ' raise event if using events
                    If m_UseEvents And (Not blnDesignTime) Then RaiseEvent MouseEnter
                End If
                ' see if need to raise events...
                If m_UseEvents Then
                    ' this horribility is one reason to not use events...
                    CopyMemory XY(0), lParam, 4&
                    ' see key states
                    Shift = Private_GetShiftMask
                    ' button states
                    If (wParam And MK_LBUTTON) = MK_LBUTTON Then Button = [Left Button]
                    If (wParam And MK_MBUTTON) = MK_MBUTTON Then Button = Button Or [Middle Button]
                    If (wParam And MK_RBUTTON) = MK_RBUTTON Then Button = Button Or [Right Button]
                    ' raise the event
                    RaiseEvent MouseMove(Button, Shift, ScaleX(XY(0), vbPixels, m_ContainerScaleMode), ScaleY(XY(1), vbPixels, m_ContainerScaleMode))
                End If
            Case WM_MOUSELEAVE
                m_MouseOver = False
                ' raise event if using events
                If m_UseEvents And (Not blnDesignTime) Then RaiseEvent MouseLeave
            Case WM_MOUSEWHEEL
                If m_UseEvents And (Not blnDesignTime) Then
                    If wParam > 0 Then
                        RaiseEvent MouseWheel([Wheel Up], Private_GetShiftMask)
                    Else
                        RaiseEvent MouseWheel([Wheel Down], Private_GetShiftMask)
                    End If
                End If
            Case WM_LBUTTONUP
                If m_UseEvents Then
                    ' this horribility is one reason to not use events...
                    CopyMemory XY(0), lParam, 4&
                    ' see key states
                    Shift = Private_GetShiftMask
                    ' button states
                    If (wParam And MK_LBUTTON) = MK_LBUTTON Then Button = [Left Button]
                    If (wParam And MK_MBUTTON) = MK_MBUTTON Then Button = Button Or [Middle Button]
                    If (wParam And MK_RBUTTON) = MK_RBUTTON Then Button = Button Or [Right Button]
                    ' raise the event
                    RaiseEvent MouseUp(Button, Shift, ScaleX(XY(0), vbPixels, m_ContainerScaleMode), ScaleY(XY(1), vbPixels, m_ContainerScaleMode))
                    ' click
                    If m_MouseOver Then RaiseEvent Click(vbLeftButton)
                End If
            Case WM_MBUTTONUP
                If m_UseEvents Then
                    ' this horribility is one reason to not use events...
                    CopyMemory XY(0), lParam, 4&
                    ' see key states
                    Shift = Private_GetShiftMask
                    ' button states
                    If (wParam And MK_LBUTTON) = MK_LBUTTON Then Button = [Left Button]
                    If (wParam And MK_MBUTTON) = MK_MBUTTON Then Button = Button Or [Middle Button]
                    If (wParam And MK_RBUTTON) = MK_RBUTTON Then Button = Button Or [Right Button]
                    ' raise the event
                    RaiseEvent MouseUp(Button, Shift, ScaleX(XY(0), vbPixels, m_ContainerScaleMode), ScaleY(XY(1), vbPixels, m_ContainerScaleMode))
                    ' click
                    If m_MouseOver Then RaiseEvent Click(vbMiddleButton)
                End If
            Case WM_RBUTTONUP
                If m_UseEvents Then
                    ' this horribility is one reason to not use events...
                    CopyMemory XY(0), lParam, 4&
                    ' see key states
                    Shift = Private_GetShiftMask
                    ' button states
                    If (wParam And MK_LBUTTON) = MK_LBUTTON Then Button = [Left Button]
                    If (wParam And MK_MBUTTON) = MK_MBUTTON Then Button = Button Or [Middle Button]
                    If (wParam And MK_RBUTTON) = MK_RBUTTON Then Button = Button Or [Right Button]
                    ' raise the event
                    RaiseEvent MouseUp(Button, Shift, ScaleX(XY(0), vbPixels, m_ContainerScaleMode), ScaleY(XY(1), vbPixels, m_ContainerScaleMode))
                    ' click
                    If m_MouseOver Then RaiseEvent Click(vbRightButton)
                End If
            Case WM_LBUTTONDBLCLK
                If m_UseEvents Then RaiseEvent DblClick(vbLeftButton)
            Case WM_MBUTTONDBLCLK
                If m_UseEvents Then RaiseEvent DblClick(vbMiddleButton)
            Case WM_RBUTTONDBLCLK
                If m_UseEvents Then RaiseEvent DblClick(vbRightButton)
        End Select
    End If
End Sub
Private Sub Private_WndProcMain(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lParamUser As Long)
    Dim XY(1) As Integer, Button As UniListMouseButton, Shift As UniListShift
    
    If bBefore Then
        Select Case uMsg
            Case WM_CTLCOLOREDIT, WM_CTLCOLORSTATIC
                SetBkMode wParam, OPAQUE
                lReturn = m_BackClrBrush
                SetBkColor wParam, m_BackClr
                SetTextColor wParam, m_ForeClr
                bHandled = True
            Case WM_DESTROY
                ssc_UnSubclass lng_hWnd
                Debug.Print "UniList: Ended subclassing! " & Hex$(lng_hWnd)
            Case WM_ERASEBKGND
                bHandled = True
                lReturn = -1&
            Case WM_COMMAND
                Select Case ((wParam And &HFFFF0000) \ &H10000) And &HFFFF&
                    'Case EN_SETFOCUS
                    '    'Debug.Print "EN_SETFOCUS"
                    '    'Private_WndEventFocus
                    '    'SetFocus m_hWnd
                    '    bHandled = True
                End Select
        End Select
        If m_UseEvents And Not bHandled Then
            Select Case uMsg
                Case WM_COMMAND
                    Select Case ((wParam And &HFFFF0000) \ &H10000) And &HFFFF&
                    End Select
            End Select
        End If
    ' BEFORE
    Else
    ' AFTER
        Select Case uMsg
            Case WM_PAINT
                If m_BorderStyle = [No Border] Then
                    Exit Sub
                ElseIf m_Appearance = [Classic 3D] Then
                    If m_BorderStyle = [Flat3D] Then
                        DrawEdge UserControl.hdc, m_RC, BDR_SUNKENOUTER, BF_RECT
                    Else
                        DrawEdge UserControl.hdc, m_RC, EDGE_SUNKEN, BF_RECT
                    End If
                End If
        End Select
    End If
End Sub

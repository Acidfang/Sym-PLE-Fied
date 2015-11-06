﻿Module mod_sendmessage_Constants

    Public Const BM_CLICK = &HF5
    Public Const BM_GETCHECK = &HF0
    Public Const SW_HIDE = 0
    Public Const SW_RESTORE = 9
    Public Const BM_SETCHECK = &HF1
    Public Const BST_UNCHECKED = &H0 ' not checked 
    Public Const BST_CHECKED = &H1 ' is checked
    Public Const WM_USER = &H400&
    Public Const CB_GETEDITSEL = &H140
    Public Const CB_LIMITTEXT = &H141
    Public Const CB_SETEDITSEL = &H142
    Public Const CB_ADDSTRING = &H143
    Public Const CB_DELETESTRING = &H144
    Public Const CB_DIR = &H145
    Public Const CB_GETCOUNT = &H146
    Public Const CB_GETCURSEL = &H147
    Public Const CB_GETLBTEXT = &H148
    Public Const CB_GETLBTEXTLEN = &H149
    Public Const CB_INSERTSTRING = &H14A
    Public Const CB_RESETCONTENT = &H14B
    Public Const CB_FINDSTRING = &H14C
    Public Const CB_SELECTSTRING = &H14D
    Public Const CB_SETCURSEL = &H14E
    Public Const CB_SHOWDROPDOWN = &H14F
    Public Const CB_GETITEMDATA = &H150
    Public Const CB_SETITEMDATA = &H151
    Public Const CB_GETDROPPEDCONTROLRECT = &H152
    Public Const CB_SETITEMHEIGHT = &H153
    Public Const CB_GETITEMHEIGHT = &H154
    Public Const CB_SETEXTENDEDUI = &H155
    Public Const CB_GETEXTENDEDUI = &H156
    Public Const CB_GETDROPPEDSTATE = &H157
    Public Const CB_FINDSTRINGEXACT = &H158
    Public Const CB_SETLOCALE = &H159
    Public Const CB_GETLOCALE = &H15A
    Public Const CB_GETTOPINDEX = &H15B
    Public Const CB_SETTOPINDEX = &H15C
    Public Const CB_GETHORIZONTALEXTENT = &H15D
    Public Const CB_SETHORIZONTALEXTENT = &H15E
    Public Const CB_GETDROPPEDWIDTH = &H15F
    Public Const CB_SETDROPPEDWIDTH = &H160
    Public Const CB_INITSTORAGE = &H161
    Public Const CB_MSGMAX = &H162
    Public Const EM_CANUNDO = &HC6
    Public Const EM_EMPTYUNDOBUFFER = &HCD
    Public Const EM_FMTLINES = &HC8
    Public Const EM_FORMATRANGE As Long = WM_USER + 57
    Public Const EM_GETFIRSTVISIBLELINE = &HCE
    Public Const EM_GETHANDLE = &HBD
    Public Const EM_GETLINE = &HC4
    Public Const EM_GETLINECOUNT = &HBA
    Public Const EM_GETMODIFY = &HB8
    Public Const EM_GETPASSWORDCHAR = &HD2
    Public Const EM_GETRECT = &HB2
    Public Const EM_GETSEL = &HB0
    Public Const EM_GETTHUMB = &HBE
    Public Const EM_GETWORDBREAKPROC = &HD1
    Public Const EM_LIMITTEXT = &HC5
    Public Const EM_LINEFROMCHAR = &HC9
    Public Const EM_LINEINDEX = &HBB
    Public Const EM_LINELENGTH = &HC1
    Public Const EM_LINESCROLL = &HB6
    Public Const EM_REPLACESEL = &HC2
    Public Const EM_SCROLL = &HB5
    Public Const EM_SCROLLCARET = &HB7
    Public Const EM_SETHANDLE = &HBC
    Public Const EM_SETMODIFY = &HB9
    Public Const EM_SETPASSWORDCHAR = &HCC
    Public Const EM_SETREADONLY = &HCF
    Public Const EM_SETRECT = &HB3
    Public Const EM_SETRECTNP = &HB4
    Public Const EM_SETSEL = &HB1
    Public Const EM_SETTABSTOPS = &HCB
    Public Const EM_SETTARGETDEVICE As Long = WM_USER + 72
    Public Const EM_SETWORDBREAKPROC = &HD0
    Public Const EM_UNDO = &HC7
    Public Const HDS_HOTTRACK = &H4
    Public Const HDI_BITMAP = &H10
    Public Const HDI_IMAGE = &H20
    Public Const HDI_ORDER = &H80
    Public Const HDI_FORMAT = &H4
    Public Const HDI_TEXT = &H2
    Public Const HDI_WIDTH = &H1
    Public Const HDI_HEIGHT = HDI_WIDTH
    Public Const HDF_LEFT = 0
    Public Const HDF_RIGHT = 1
    Public Const HDF_IMAGE = &H800
    Public Const HDF_BITMAP_ON_RIGHT = &H1000
    Public Const HDF_BITMAP = &H2000
    Public Const HDF_STRING = &H4000
    Public Const HDM_FIRST = &H1200
    Public Const HDM_SETITEM = (HDM_FIRST + 4)
    Public Const LB_ADDFILE = &H196
    Public Const LB_ADDSTRING = &H180
    Public Const LB_CTLCODE = 0&
    Public Const LB_DELETESTRING = &H182
    Public Const LB_DIR = &H18D
    Public Const LB_ERR = (-1)
    Public Const LB_ERRSPACE = (-2)
    Public Const LB_FINDSTRING = &H18F
    Public Const LB_FINDSTRINGEXACT = &H1A2
    Public Const LB_GETANCHORINDEX = &H19D
    Public Const LB_GETCARETINDEX = &H19F
    Public Const LB_GETCOUNT = &H18B
    Public Const LB_GETCURSEL = &H188
    Public Const LB_GETHORIZONTALEXTENT = &H193
    Public Const LB_GETITEMDATA = &H199
    Public Const LB_GETITEMHEIGHT = &H1A1
    Public Const LB_GETITEMRECT = &H198
    Public Const LB_GETLOCALE = &H1A6
    Public Const LB_GETSEL = &H187
    Public Const LB_GETSELCOUNT = &H190
    Public Const LB_GETSELITEMS = &H191
    Public Const LB_GETTEXT = &H189
    Public Const LB_GETTEXTLEN = &H18A
    Public Const LB_GETTOPINDEX = &H18E
    Public Const LB_INSERTSTRING = &H181
    Public Const LB_MSGMAX = &H1A8
    Public Const LB_OKAY = 0
    Public Const LB_RESETCONTENT = &H184
    Public Const LB_SELECTSTRING = &H18C
    Public Const LB_SELITEMRANGE = &H19B
    Public Const LB_SELITEMRANGEEX = &H183
    Public Const LB_SETANCHORINDEX = &H19C
    Public Const LB_SETCARETINDEX = &H19E
    Public Const LB_SETCOLUMNWIDTH = &H195
    Public Const LB_SETCOUNT = &H1A7
    Public Const LB_SETCURSEL = &H186
    Public Const LB_SETHORIZONTALEXTENT = &H194
    Public Const LB_SETITEMDATA = &H19A
    Public Const LB_SETITEMHEIGHT = &H1A0
    Public Const LB_SETLOCALE = &H1A5
    Public Const LB_SETSEL = &H185
    Public Const LB_SETTABSTOPS = &H192
    Public Const LB_SETTOPINDEX = &H197
    Public Const LBN_DBLCLK = 2
    Public Const LBN_ERRSPACE = (-2)
    Public Const LBN_KILLFOCUS = 5
    Public Const LBN_SELCANCEL = 3
    Public Const LBN_SELCHANGE = 1
    Public Const LBN_SETFOCUS = 4
    Public Const LVM_FIRST = &H1000
    Public Const LVM_GETHEADER = (LVM_FIRST + 31)
    Public Const LVM_GETBKCOLOR = (LVM_FIRST + 0)
    Public Const LVM_SETBKCOLOR = (LVM_FIRST + 1)
    Public Const LVM_GETIMAGELIST = (LVM_FIRST + 2)
    Public Const LVM_SETIMAGELIST = (LVM_FIRST + 3)
    Public Const LVM_GETITEMCOUNT = (LVM_FIRST + 4)
    Public Const LVM_GETITEMA = (LVM_FIRST + 5)
    Public Const LVM_GETITEM = LVM_GETITEMA
    Public Const LVM_SETITEMA = (LVM_FIRST + 6)
    Public Const LVM_SETITEM = LVM_SETITEMA
    Public Const LVM_INSERTITEMA = (LVM_FIRST + 7)
    Public Const LVM_INSERTITEM = LVM_INSERTITEMA
    Public Const LVM_DELETEITEM = (LVM_FIRST + 8)
    Public Const LVM_DELETEALLITEMS = (LVM_FIRST + 9)
    Public Const LVM_GETCALLBACKMASK = (LVM_FIRST + 10)
    Public Const LVM_SETCALLBACKMASK = (LVM_FIRST + 11)
    Public Const LVM_GETNEXTITEM = (LVM_FIRST + 12)
    Public Const LVM_FINDITEMA = (LVM_FIRST + 13)
    Public Const LVM_FINDITEM = LVM_FINDITEMA
    Public Const LVM_GETITEMRECT = (LVM_FIRST + 14)
    Public Const LVM_SETITEMPOSITION = (LVM_FIRST + 15)
    Public Const LVM_GETITEMPOSITION = (LVM_FIRST + 16)
    Public Const LVM_GETSTRINGWIDTHA = (LVM_FIRST + 17)
    Public Const LVM_GETSTRINGWIDTH = LVM_GETSTRINGWIDTHA
    Public Const LVM_HITTEST = (LVM_FIRST + 18)
    Public Const LVM_ENSUREVISIBLE = (LVM_FIRST + 19)
    Public Const LVM_SCROLL = (LVM_FIRST + 20)
    Public Const LVM_REDRAWITEMS = (LVM_FIRST + 21)
    Public Const LVM_ARRANGE = (LVM_FIRST + 22)
    Public Const LVM_EDITLABELA = (LVM_FIRST + 23)
    Public Const LVM_EDITLABEL = LVM_EDITLABELA
    Public Const LVM_GETEDITCONTROL = (LVM_FIRST + 24)
    Public Const LVM_GETCOLUMNA = (LVM_FIRST + 25)
    Public Const LVM_GETCOLUMN = LVM_GETCOLUMNA
    Public Const LVM_SETCOLUMNA = (LVM_FIRST + 26)
    Public Const LVM_SETCOLUMN = LVM_SETCOLUMNA
    Public Const LVM_INSERTCOLUMNA = (LVM_FIRST + 27)
    Public Const LVM_INSERTCOLUMN = LVM_INSERTCOLUMNA
    Public Const LVM_DELETECOLUMN = (LVM_FIRST + 28)
    Public Const LVM_GETCOLUMNWIDTH = (LVM_FIRST + 29)
    Public Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
    Public Const LVM_CREATEDRAGIMAGE = (LVM_FIRST + 33)
    Public Const LVM_GETVIEWRECT = (LVM_FIRST + 34)
    Public Const LVM_GETTEXTCOLOR = (LVM_FIRST + 35)
    Public Const LVM_SETTEXTCOLOR = (LVM_FIRST + 36)
    Public Const LVM_GETTEXTBKCOLOR = (LVM_FIRST + 37)
    Public Const LVM_SETTEXTBKCOLOR = (LVM_FIRST + 38)
    Public Const LVM_GETTOPINDEX = (LVM_FIRST + 39)
    Public Const LVM_GETCOUNTPERPAGE = (LVM_FIRST + 40)
    Public Const LVM_GETORIGIN = (LVM_FIRST + 41)
    Public Const LVM_UPDATE = (LVM_FIRST + 42)
    Public Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
    Public Const LVM_GETITEMSTATE = (LVM_FIRST + 44)
    Public Const LVM_GETITEMTEXTA = (LVM_FIRST + 45)
    Public Const LVM_GETITEMTEXT = LVM_GETITEMTEXTA
    Public Const LVM_SETITEMTEXTA = (LVM_FIRST + 46)
    Public Const LVM_SETITEMTEXT = LVM_SETITEMTEXTA
    Public Const LVM_SETITEMCOUNT = (LVM_FIRST + 47)
    Public Const LVM_SORTITEMS = (LVM_FIRST + 48)
    Public Const LVM_SETITEMPOSITION32 = (LVM_FIRST + 49)
    Public Const LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)
    Public Const LVM_GETITEMSPACING = (LVM_FIRST + 51)
    Public Const LVM_GETISEARCHSTRINGA = (LVM_FIRST + 52)
    Public Const LVM_GETISEARCHSTRING = LVM_GETISEARCHSTRINGA
    Public Const LVM_SETICONSPACING = (LVM_FIRST + 53)
    Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54)
    Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55)
    Public Const LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
    Public Const LVM_SUBITEMHITTEST = (LVM_FIRST + 57)
    Public Const LVM_SETCOLUMNORDERARRAY = (LVM_FIRST + 58)
    Public Const LVM_GETCOLUMNORDERARRAY = (LVM_FIRST + 59)
    Public Const LVM_SETHOTITEM = (LVM_FIRST + 60)
    Public Const LVM_GETHOTITEM = (LVM_FIRST + 61)
    Public Const LVM_SETHOTCURSOR = (LVM_FIRST + 62)
    Public Const LVM_GETHOTCURSOR = (LVM_FIRST + 63)
    Public Const LVM_APPROXIMATEVIEWRECT = (LVM_FIRST + 64)
    Public Const LVS_EX_FULLROWSELECT = &H20
    Public Const LVSCW_AUTOSIZE As Long = -1
    Public Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
    Public Const WM_ACTIVATE = &H6
    Public Const WM_ACTIVATEAPP = &H1C
    Public Const WM_ASKCBFORMATNAME = &H30C
    Public Const WM_CANCELJOURNAL = &H4B
    Public Const WM_CANCELMODE = &H1F
    Public Const WM_CHANGECBCHAIN = &H30D
    Public Const WM_CHAR = &H102
    Public Const WM_CHARTOITEM = &H2F
    Public Const WM_CHILDACTIVATE = &H22
    Public Const WM_CHOOSEFONT_GETLOGFONT = (WM_USER + 1)
    Public Const WM_CHOOSEFONT_SETFLAGS = (WM_USER + 102)
    Public Const WM_CHOOSEFONT_SETLOGFONT = (WM_USER + 101)
    Public Const WM_CLEAR = &H303
    Public Const WM_CLOSE = &H10
    Public Const WM_COMMAND = &H111
    Public Const WM_COMMNOTIFY = &H44 ' no longer suported
    Public Const WM_COMPACTING = &H41
    Public Const WM_COMPAREITEM = &H39
    Public Const WM_CONVERTREQUESTEX = &H108
    Public Const WM_COPY = &H301
    Public Const WM_COPYDATA = &H4A
    Public Const WM_CREATE = &H1
    Public Const WM_CTLCOLORBTN = &H135
    Public Const WM_CTLCOLORDLG = &H136
    Public Const WM_CTLCOLOREDIT = &H133
    Public Const WM_CTLCOLORLISTBOX = &H134
    Public Const WM_CTLCOLORMSGBOX = &H132
    Public Const WM_CTLCOLORSCROLLBAR = &H137
    Public Const WM_CTLCOLORSTATIC = &H138
    Public Const WM_CUT = &H300
    Public Const WM_DDE_FIRST = &H3E0
    Public Const WM_DDE_ACK = (WM_DDE_FIRST + 4)
    Public Const WM_DDE_ADVISE = (WM_DDE_FIRST + 2)
    Public Const WM_DDE_DATA = (WM_DDE_FIRST + 5)
    Public Const WM_DDE_EXECUTE = (WM_DDE_FIRST + 8)
    Public Const WM_DDE_INITIATE = (WM_DDE_FIRST)
    Public Const WM_DDE_LAST = (WM_DDE_FIRST + 8)
    Public Const WM_DDE_POKE = (WM_DDE_FIRST + 7)
    Public Const WM_DDE_REQUEST = (WM_DDE_FIRST + 6)
    Public Const WM_DDE_TERMINATE = (WM_DDE_FIRST + 1)
    Public Const WM_DDE_UNADVISE = (WM_DDE_FIRST + 3)
    Public Const WM_DEADCHAR = &H103
    Public Const WM_DELETEITEM = &H2D
    Public Const WM_DESTROY = &H2
    Public Const WM_DESTROYCLIPBOARD = &H307
    Public Const WM_DEVMODECHANGE = &H1B
    Public Const WM_DRAWCLIPBOARD = &H308
    Public Const WM_DRAWITEM = &H2B
    Public Const WM_DROPFILES = &H233
    Public Const WM_ENABLE = &HA
    Public Const WM_ENDSESSION = &H16
    Public Const WM_ENTERIDLE = &H121
    Public Const WM_ENTERMENULOOP = &H211
    Public Const WM_ERASEBKGND = &H14
    Public Const WM_EXITMENULOOP = &H212
    Public Const WM_FONTCHANGE = &H1D
    Public Const WM_GETFONT = &H31
    Public Const WM_GETDLGCODE = &H87
    Public Const WM_GETHOTKEY = &H33
    Public Const WM_GETMINMAXINFO = &H24
    Public Const WM_GETTEXT = &HD
    Public Const WM_GETTEXTLENGTH = &HE
    Public Const WM_HOTKEY = &H312
    Public Const WM_HSCROLL = &H114
    Public Const WM_HSCROLLCLIPBOARD = &H30E
    Public Const WM_ICONERASEBKGND = &H27
    Public Const WM_IME_CHAR = &H286
    Public Const WM_IME_COMPOSITION = &H10F
    Public Const WM_IME_COMPOSITIONFULL = &H284
    Public Const WM_IME_CONTROL = &H283
    Public Const WM_IME_ENDCOMPOSITION = &H10E
    Public Const WM_IME_KEYDOWN = &H290
    Public Const WM_IME_KEYLAST = &H10F
    Public Const WM_IME_KEYUP = &H291
    Public Const WM_IME_NOTIFY = &H282
    Public Const WM_IME_SELECT = &H285
    Public Const WM_IME_SETCONTEXT = &H281
    Public Const WM_IME_STARTCOMPOSITION = &H10D
    Public Const WM_INITDIALOG = &H110
    Public Const WM_INITMENU = &H116
    Public Const WM_INITMENUPOPUP = &H117
    Public Const WM_KEYDOWN = &H100
    Public Const WM_KEYFIRST = &H100
    Public Const WM_KEYLAST = &H108
    Public Const WM_KEYUP = &H101
    Public Const WM_KILLFOCUS = &H8
    Public Const WM_LBUTTONDBLCLK = &H203
    Public Const WM_LBUTTONDOWN = &H201
    Public Const WM_LBUTTONUP = &H202
    Public Const WM_MBUTTONDBLCLK = &H209
    Public Const WM_MBUTTONDOWN = &H207
    Public Const WM_MBUTTONUP = &H208
    Public Const WM_MDIACTIVATE = &H222
    Public Const WM_MDICASCADE = &H227
    Public Const WM_MDICREATE = &H220
    Public Const WM_MDIDESTROY = &H221
    Public Const WM_MDIGETACTIVE = &H229
    Public Const WM_MDIICONARRANGE = &H228
    Public Const WM_MDIMAXIMIZE = &H225
    Public Const WM_MDINEXT = &H224
    Public Const WM_MDIREFRESHMENU = &H234
    Public Const WM_MDIRESTORE = &H223
    Public Const WM_MDISETMENU = &H230
    Public Const WM_MDITILE = &H226
    Public Const WM_MEASUREITEM = &H2C
    Public Const WM_MENUCHAR = &H120
    Public Const WM_MENUSELECT = &H11F
    Public Const WM_MOUSEACTIVATE = &H21
    Public Const WM_MOUSEFIRST = &H200
    Public Const WM_MOUSELAST = &H209
    Public Const WM_MOUSEMOVE = &H200
    Public Const WM_MOVE = &H3
    Public Const WM_NCACTIVATE = &H86
    Public Const WM_NCCALCSIZE = &H83
    Public Const WM_NCCREATE = &H81
    Public Const WM_NCDESTROY = &H82
    Public Const WM_NCHITTEST = &H84
    Public Const WM_NCLBUTTONDBLCLK = &HA3
    Public Const WM_NCLBUTTONDOWN = &HA1
    Public Const WM_NCLBUTTONUP = &HA2
    Public Const WM_NCMBUTTONDBLCLK = &HA9
    Public Const WM_NCMBUTTONDOWN = &HA7
    Public Const WM_NCMBUTTONUP = &HA8
    Public Const WM_NCMOUSEMOVE = &HA0
    Public Const WM_NCPAINT = &H85
    Public Const WM_NCRBUTTONDBLCLK = &HA6
    Public Const WM_NCRBUTTONDOWN = &HA4
    Public Const WM_NCRBUTTONUP = &HA5
    Public Const WM_NEXTDLGCTL = &H28
    Public Const WM_NULL = &H0
    Public Const WM_OTHERWINDOWCREATED = &H42 ' no longer suported
    Public Const WM_OTHERWINDOWDESTROYED = &H43 ' no longer suported
    Public Const WM_PAINT = &HF
    Public Const WM_PAINTCLIPBOARD = &H309
    Public Const WM_PAINTICON = &H26
    Public Const WM_PALETTECHANGED = &H311
    Public Const WM_PALETTEISCHANGING = &H310
    Public Const WM_PARENTNOTIFY = &H210
    Public Const WM_PASTE = &H302
    Public Const WM_PENWINFIRST = &H380
    Public Const WM_PENWINLAST = &H38F
    Public Const WM_POWER = &H48
    Public Const WM_PSD_ENVSTAMPRECT = (WM_USER + 5)
    Public Const WM_PSD_FULLPAGERECT = (WM_USER + 1)
    Public Const WM_PSD_GREEKTEXTRECT = (WM_USER + 4)
    Public Const WM_PSD_MARGINRECT = (WM_USER + 3)
    Public Const WM_PSD_MINMARGINRECT = (WM_USER + 2)
    Public Const WM_PSD_PAGESETUPDLG = (WM_USER)
    Public Const WM_PSD_YAFULLPAGERECT = (WM_USER + 6)
    Public Const WM_QUERYDRAGICON = &H37
    Public Const WM_QUERYENDSESSION = &H11
    Public Const WM_QUERYNEWPALETTE = &H30F
    Public Const WM_QUERYOPEN = &H13
    Public Const WM_QUEUESYNC = &H23
    Public Const WM_QUIT = &H12
    Public Const WM_RBUTTONDBLCLK = &H206
    Public Const WM_RBUTTONDOWN = &H204
    Public Const WM_RBUTTONUP = &H205
    Public Const WM_RENDERALLFORMATS = &H306
    Public Const WM_RENDERFORMAT = &H305
    Public Const WM_SETCURSOR = &H20
    Public Const WM_SETFOCUS = &H7
    Public Const WM_SETFONT = &H30
    Public Const WM_SETHOTKEY = &H32
    Public Const WM_SETREDRAW = &HB
    Public Const WM_SETTEXT = &HC
    Public Const WM_SHOWWINDOW = &H18
    Public Const WM_SIZE = &H5
    Public Const WM_SIZECLIPBOARD = &H30B
    Public Const WM_SPOOLERSTATUS = &H2A
    Public Const WM_SYSCHAR = &H106
    Public Const WM_SYSCOLORCHANGE = &H15
    Public Const WM_SYSCOMMAND = &H112
    Public Const WM_SYSDEADCHAR = &H107
    Public Const WM_SYSKEYDOWN = &H104
    Public Const WM_SYSKEYUP = &H105
    Public Const WM_TIMECHANGE = &H1E
    Public Const WM_TIMER = &H113
    Public Const WM_UNDO = &H304
    Public Const WM_VKEYTOITEM = &H2E
    Public Const WM_VSCROLL = &H115
    Public Const WM_VSCROLLCLIPBOARD = &H30A
    Public Const WM_WINDOWPOSCHANGED = &H47
    Public Const WM_WINDOWPOSCHANGING = &H46
    Public Const WM_WININICHANGE = &H1A
    Public Const WS_BORDER = &H800000
    Public Const WS_CAPTION = &HC00000 ' WS_BORDER Or WS_DLGFRAME
    Public Const WS_CHILD = &H40000000
    Public Const WS_CHILDWINDOW = (WS_CHILD)
    Public Const WS_CLIPCHILDREN = &H2000000
    Public Const WS_CLIPSIBLINGS = &H4000000
    Public Const WS_DISABLED = &H8000000
    Public Const WS_DLGFRAME = &H400000
    Public Const WS_EX_ACCEPTFILES = &H10&
    Public Const WS_EX_DLGMODALFRAME = &H1&
    Public Const WS_EX_NOPARENTNOTIFY = &H4&
    Public Const WS_EX_TOPMOST = &H8&
    Public Const WS_EX_TRANSPARENT = &H20&
    Public Const WS_GROUP = &H20000
    Public Const WS_HSCROLL = &H100000
    Public Const WS_MINIMIZE = &H20000000
    Public Const WS_ICONIC = WS_MINIMIZE
    Public Const WS_MAXIMIZE = &H1000000
    Public Const WS_MAXIMIZEBOX = &H10000
    Public Const WS_MINIMIZEBOX = &H20000
    Public Const WS_SYSMENU = &H80000
    Public Const WS_THICKFRAME = &H40000
    Public Const WS_OVERLAPPED = &H0&
    Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION _
    Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
    Public Const WS_POPUP = &H80000000
    Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
    Public Const WS_SIZEBOX = WS_THICKFRAME
    Public Const WS_TABSTOP = &H10000
    Public Const WS_TILED = WS_OVERLAPPED
    Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
    Public Const WS_VISIBLE = &H10000000
    Public Const WS_VSCROLL = &H200000
    Public Const LBS_DISABLENOSCROLL = &H1000&
    Public Const LBS_EXTENDEDSEL = &H800&
    Public Const LBS_HASSTRINGS = &H40&
    Public Const LBS_MULTICOLUMN = &H200&
    Public Const LBS_MULTIPLESEL = &H8&
    Public Const LBS_NODATA = &H2000&
    Public Const LBS_NOINTEGRALHEIGHT = &H100&
    Public Const LBS_NOREDRAW = &H4&
    Public Const LBS_NOTIFY = &H1&
    Public Const LBS_OWNERDRAWFIXED = &H10&
    Public Const LBS_OWNERDRAWVARIABLE = &H20&
    Public Const LBS_SORT = &H2&
    Public Const LBS_STANDARD = (LBS_NOTIFY Or LBS_SORT Or WS_VSCROLL _
    Or WS_BORDER)
    Public Const LBS_USETABSTOPS = &H80&
    Public Const LBS_WANTKEYBOARDINPUT = &H400&
    Public Const LBSELCHSTRING = "commdlg_LBSelChangedNotify"
    Public Const TB_ENABLEBUTTON = (WM_USER + 1)
    Public Const TB_CHECKBUTTON = (WM_USER + 2)
    Public Const TB_PRESSBUTTON = (WM_USER + 3)
    Public Const TB_HIDEBUTTON = (WM_USER + 4)
    Public Const TB_INDETERMINATE = (WM_USER + 5)
    Public Const TB_MARKBUTTON = (WM_USER + 6)
    Public Const TB_ISBUTTONENABLED = (WM_USER + 9)
    Public Const TB_ISBUTTONCHECKED = (WM_USER + 10)
    Public Const TB_ISBUTTONPRESSED = (WM_USER + 11)
    Public Const TB_ISBUTTONHIDDEN = (WM_USER + 12)
    Public Const TB_ISBUTTONINDETERMINATE = (WM_USER + 13)
    Public Const TB_ISBUTTONHIGHLIGHTED = (WM_USER + 14)
    Public Const TB_SETSTATE = (WM_USER + 17)
    Public Const TB_GETSTATE = (WM_USER + 18)
    Public Const TB_ADDBITMAP = (WM_USER + 19)
    Public Const TB_ADDBUTTONSA = (WM_USER + 20)
    Public Const TB_INSERTBUTTONA = (WM_USER + 21)
    Public Const TB_ADDBUTTONS = (WM_USER + 20)
    Public Const TB_INSERTBUTTON = (WM_USER + 21)
    Public Const TB_DELETEBUTTON = (WM_USER + 22)
    Public Const TB_GETBUTTON = (WM_USER + 23)
    Public Const TB_BUTTONCOUNT = (WM_USER + 24)
    Public Const TB_COMMANDTOINDEX = (WM_USER + 25)
    Public Const TB_SAVERESTOREA = (WM_USER + 26)
    Public Const TB_SAVERESTOREW = (WM_USER + 76)
    Public Const TB_CUSTOMIZE = (WM_USER + 27)
    Public Const TB_ADDSTRINGA = (WM_USER + 28)
    Public Const TB_ADDSTRINGW = (WM_USER + 77)
    Public Const TB_GETITEMRECT = (WM_USER + 29)
    Public Const TB_BUTTONSTRUCTSIZE = (WM_USER + 30)
    Public Const TB_SETBUTTONSIZE = (WM_USER + 31)
    Public Const TB_SETBITMAPSIZE = (WM_USER + 32)
    Public Const TB_AUTOSIZE = (WM_USER + 33)
    Public Const TB_GETTOOLTIPS = (WM_USER + 35)
    Public Const TB_SETTOOLTIPS = (WM_USER + 36)
    Public Const TB_SETPARENT = (WM_USER + 37)
    Public Const TB_SETROWS = (WM_USER + 39)
    Public Const TB_GETROWS = (WM_USER + 40)
    Public Const TB_SETCMDID = (WM_USER + 42)
    Public Const TB_CHANGEBITMAP = (WM_USER + 43)
    Public Const TB_GETBITMAP = (WM_USER + 44)
    Public Const TB_GETBUTTONTEXTA = (WM_USER + 45)
    Public Const TB_GETBUTTONTEXTW = (WM_USER + 75)
    Public Const TB_REPLACEBITMAP = (WM_USER + 46)
    Public Const TB_SETINDENT = (WM_USER + 47)
    Public Const TB_SETIMAGELIST = (WM_USER + 48)
    Public Const TB_GETIMAGELIST = (WM_USER + 49)
    Public Const TB_LOADIMAGES = (WM_USER + 50)
    Public Const TB_GETRECT = (WM_USER + 51) ' wParam is the Cmd instead of index
    Public Const TB_SETHOTIMAGELIST = (WM_USER + 52)
    Public Const TB_GETHOTIMAGELIST = (WM_USER + 53)
    Public Const TB_SETDISABLEDIMAGELIST = (WM_USER + 54)
    Public Const TB_GETDISABLEDIMAGELIST = (WM_USER + 55)
    Public Const TB_SETSTYLE = (WM_USER + 56)
    Public Const TB_GETSTYLE = (WM_USER + 57)
    Public Const TB_GETBUTTONSIZE = (WM_USER + 58)
    Public Const TB_SETBUTTONWIDTH = (WM_USER + 59)
    Public Const TB_SETMAXTEXTROWS = (WM_USER + 60)
    Public Const TB_GETTEXTROWS = (WM_USER + 61)
    Public Const TBSTYLE_BUTTON = &H0
    Public Const TBSTYLE_SEP = &H1
    Public Const TBSTYLE_CHECK = &H2
    Public Const TBSTYLE_GROUP = &H4
    Public Const TBSTYLE_CHECKGROUP = (TBSTYLE_GROUP Or TBSTYLE_CHECK)
    Public Const TBSTYLE_DROPDOWN = &H8
    Public Const TBSTYLE_AUTOSIZE = &H10 ' automatically calculate the cx of the button '
    Public Const TBSTYLE_NOPREFIX = &H20 ' If this button should not have accel prefix '
    Public Const TBSTYLE_TOOLTIPS = &H100
    Public Const TBSTYLE_WRAPABLE = &H200
    Public Const TBSTYLE_ALTDRAG = &H400
    Public Const TBSTYLE_FLAT = &H800
    Public Const TBSTYLE_LIST = &H1000
    Public Const TBSTYLE_CUSTOMERASE = &H2000
    Public Const TBSTYLE_REGISTERDROP = &H4000
    Public Const TBSTYLE_TRANSPARENT = &H8000
    Public Const TBSTYLE_EX_DRAWDDARROWS = &H1
    Public Const VK_LBUTTON = &H1
    Public Const VK_RBUTTON = &H2
    Public Const VK_CANCEL = &H3
    Public Const VK_MBUTTON = &H4 ' NOT contiguous with L RBUTTON
    Public Const VK_BACK = &H8
    Public Const VK_TAB = &H9
    Public Const VK_CLEAR = &HC
    Public Const VK_RETURN = &HD
    Public Const VK_SHIFT = &H10
    Public Const VK_CONTROL = &H11
    Public Const VK_MENU = &H12
    Public Const VK_PAUSE = &H13
    Public Const VK_CAPITAL = &H14
    Public Const VK_ESCAPE = &H1B
    Public Const VK_SPACE = &H20
    Public Const VK_PRIOR = &H21
    Public Const VK_NEXT = &H22
    Public Const VK_END = &H23
    Public Const VK_HOME = &H24
    Public Const VK_LEFT = &H25
    Public Const VK_UP = &H26
    Public Const VK_RIGHT = &H27
    Public Const VK_DOWN = &H28
    Public Const VK_SELECT = &H29
    Public Const VK_PRINT = &H2A
    Public Const VK_EXECUTE = &H2B
    Public Const VK_SNAPSHOT = &H2C
    Public Const VK_INSERT = &H2D
    Public Const VK_DELETE = &H2E
    Public Const VK_HELP = &H2F
    Public Const KEYEVENTF_EXTENDEDKEY As UInteger = &H1
    Public Const KEYEVENTF_KEYUP As UInteger = &H2
    Public Const GW_CHILD As Integer = 5
    Public Const GW_HWNDNEXT As Integer = 2
End Module
HA$PBExportHeader$cxls.sru
forward
global type cxls from nonvisualobject
end type
end forward

global type cxls from nonvisualobject
end type
global cxls cxls

type variables
//File Type
constant int TYPE_XLS = 0
constant int TYPE_XLSX	=	1

//cell type
constant long CELLTYPE_EMPTY 	= 0
constant long CELLTYPE_NUMBER = 1
constant long CELLTYPE_STRING = 2
constant long CELLTYPE_BOOLEAN 	= 3
constant long CELLTYPE_BLANK 	= 4
constant long CELLTYPE_ERROR 	= 5
constant long CELLTYPE_DATETIME 	= 6

//alignh
constant long ALIGNH_GENERAL 	=	0
constant long ALIGNH_LEFT		=	1
constant long ALIGNH_CENTER	=	2
constant long ALIGNH_RIGHT		=	3
constant long ALIGNH_FILL		=	4
constant long ALIGNH_JUSTIFY	=	5
constant long ALIGNH_MERGE		=	6
constant long ALIGNH_DISTRIBUTED	=	7

//alignv
constant long ALIGNV_TOP		=	0
constant long ALIGNV_CENTER	=	1
constant long ALIGNV_BOTTOM	=	2
constant long ALIGNV_JUSTIFY	=	3
constant long ALIGNV_DISTRIBUTED	=	4

//borderstyle
constant long  BORDERSTYLE_NONE				=	0	
constant long  BORDERSTYLE_THIN				=	1
constant long  BORDERSTYLE_MEDIUM			=	2
constant long  BORDERSTYLE_DASHED			=	3
constant long  BORDERSTYLE_DOTTED			=	4
constant long  BORDERSTYLE_THICK				=	5
constant long  BORDERSTYLE_DOUBLE			=	6
constant long  BORDERSTYLE_HAIR				=	7
constant long  BORDERSTYLE_MEDIUMDASHED	=	8
constant long  BORDERSTYLE_DASHDOT			=	9
constant long  BORDERSTYLE_MEDIUMDASHDOT	=	10
constant long  BORDERSTYLE_DASHDOTDOT		=	11
constant long  BORDERSTYLE_MEDIUMDASHDOTDOT		=	12
constant long  BORDERSTYLE_SLANTDASHDOT			=	13

//numformat
constant long NUMFORMAT_GENERAL	=	0					//general format	
constant long NUMFORMAT_NUMBER	=	1					//general number	1000
constant long NUMFORMAT_NUMBER_D2	=	2				//number with decimal point	1000.00
constant long NUMFORMAT_NUMBER_SEP	=	3				//number with thousands separator	100000
constant long NUMFORMAT_NUMBER_SEP_D2	=	4			//number with decimal point and thousands separator	100000.00
constant long NUMFORMAT_CURRENCY_NEGBRA	=	5		//monetary value negative in brackets	(1000$)
constant long NUMFORMAT_CURRENCY_NEGBRARED	=	6	//monetary value negative is red in brackets	(1000$)
constant long NUMFORMAT_CURRENCY_D2_NEGBRA	=	7	//monetary value with decimal point negative in brackets	($1000.00)
constant long NUMFORMAT_CURRENCY_D2_NEGBRARED	=	8	//monetary value with decimal point negative is red in brackets	($1000.00)
constant long NUMFORMAT_PERCENT	=	9					//percent value multiply the cell value by 100	75%
constant long NUMFORMAT_PERCENT_D2	=	10				//percent value with decimal point multiply the cell value by 100	75.00%
constant long NUMFORMAT_SCIENTIFIC_D2	=	11			//scientific value with E character and decimal point	10.00E+1
constant long NUMFORMAT_FRACTION_ONEDIG	=	12		//fraction value one digit	10 1/2
constant long NUMFORMAT_FRACTION_TWODIG	=	13		//fraction value two digits	10 23/95
constant long NUMFORMAT_DATE						=	14	//date value depends on OS settings	3/11/2009
constant long NUMFORMAT_CUSTOM_D_MON_YY		=	15	//custom date value	11-Mar-09
constant long NUMFORMAT_CUSTOM_D_MON			=	16	//custom date value	11-Mar
constant long NUMFORMAT_CUSTOM_MON_YY			=	17	//custom date value	Mar-09
constant long NUMFORMAT_CUSTOM_HMM_AM			=	18	//custom date value	8:30 AM
constant long NUMFORMAT_CUSTOM_HMMSS_AM		=	19	//custom date value	8:30:00 AM
constant long NUMFORMAT_CUSTOM_HMM				=	20	//custom date value	8:30
constant long NUMFORMAT_CUSTOM_HMMSS			=	21	//custom date value	8:30:00
constant long NUMFORMAT_CUSTOM_MDYYYY_HMM			=	22//custom datetime value	3/11/2009 8:30
constant long NUMFORMAT_NUMBER_SEP_NEGBRA			=	37//number with thousands separator negative in brackets	(4000)
constant long NUMFORMAT_NUMBER_SEP_NEGBRARED		=	38	//number with thousands separator negative is red in brackets	(4000)
constant long NUMFORMAT_NUMBER_D2_SEP_NEGBRA		=	39	//number with thousands separator and decimal point negative in brackets	(4000.00)
constant long NUMFORMAT_NUMBER_D2_SEP_NEGBRARED	=	40	//number with thousands separator and decimal point negative is red in brackets	(4000.00)
constant long NUMFORMAT_ACCOUNT		=	41					//account value	5000
constant long NUMFORMAT_ACCOUNTCUR		=	42				//account value with currency symbol	$	5000
constant long NUMFORMAT_ACCOUNT_D2		=	43				//account value with decimal point	5000.00
constant long NUMFORMAT_ACCOUNT_D2_CUR		=	44			//account value with currency symbol and decimal point	$	5000.00
constant long NUMFORMAT_CUSTOM_MMSS		=	45				//custom time value	30:55
constant long NUMFORMAT_CUSTOM_H0MMSS		=	46			//custom time value	20:30:55
constant long NUMFORMAT_CUSTOM_MMSS0		=	47			//custom time value	30:55.0
constant long NUMFORMAT_CUSTOM_000P0E_PLUS0		=	48	//custom value	15.2E+3
constant long NUMFORMAT_TEXT		=	49 					//text value	any text

//color
constant long COLOR_BLACK = 8
constant long COLOR_WHITE =  9
constant long COLOR_RED =  10
constant long COLOR_BRIGHTGREEN =	11  
constant long COLOR_BLUE =  	12
constant long COLOR_YELLOW =  13
constant long COLOR_PINK = 14
constant long COLOR_TURQUOISE =	15  
constant long COLOR_DARKRED =     16     
constant long COLOR_GREEN =  	17
constant long COLOR_DARKBLUE =  18
constant long COLOR_DARKYELLOW =  19
constant long COLOR_VIOLET =  	20
constant long COLOR_TEAL =  	21
constant long COLOR_GRAY25 =  	22
constant long COLOR_GRAY50 =  	23
constant long COLOR_PERIWINKLE_CF =	24                
constant long COLOR_PLUM_CF =  		25
constant long COLOR_IVORY_CF =  	26
constant long COLOR_LIGHTTURQUOISE_CF =  27
constant long COLOR_DARKPURPLE_CF =  	28
constant long COLOR_CORAL_CF =  	29
constant long COLOR_OCEANBLUE_CF =  	30
constant long COLOR_ICEBLUE_CF =          	31      
constant long COLOR_DARKBLUE_CL =  	32
constant long COLOR_PINK_CL =  		33
constant long COLOR_YELLOW_CL =  	34
constant long COLOR_TURQUOISE_CL =  	35
constant long COLOR_VIOLET_CL =  	36
constant long COLOR_DARKRED_CL =  	37
constant long COLOR_TEAL_CL =             38   
constant long COLOR_BLUE_CL =  		39
constant long COLOR_SKYBLUE =  		40
constant long COLOR_LIGHTTURQUOISE =  	41
constant long COLOR_LIGHTGREEN =  	42
constant long COLOR_LIGHTYELLOW =  	43
constant long COLOR_PALEBLUE =  	44
constant long COLOR_ROSE =  		45
constant long COLOR_LAVENDER =           46  
constant long COLOR_TAN =  		47
constant long COLOR_LIGHTBLUE =  	48
constant long COLOR_AQUA =  		49
constant long COLOR_LIME =  		50
constant long COLOR_GOLD =  		51
constant long COLOR_LIGHTORANGE =  52
constant long COLOR_ORANGE =  53
constant long COLOR_BLUEGRAY =  54
constant long COLOR_GRAY40 =       55     
constant long COLOR_DARKTEAL =  56
constant long COLOR_SEAGREEN =  57
constant long COLOR_DARKGREEN =  58
constant long COLOR_OLIVEGREEN =  59
constant long COLOR_BROWN =  60
constant long COLOR_PLUM =  61
constant long COLOR_INDIGO =  62
constant long COLOR_GRAY80 =     63      
constant long COLOR_DEFAULT_FOREGROUND = 64	// 0x0040 =  64
constant long COLOR_DEFAULT_BACKGROUND = 65	//0x0041 =  65
constant long COLOR_TOOLTIP = 81	//0x0051 
constant long COLOR_NONE = 	127	//0x7F =  
constant long COLOR_AUTO = 	32767	//0x7FFF


constant long UNDERLINE_NONE	=	0
constant long UNDERLINE_SINGLE	=	1
constant long UNDERLINE_DOUBLE	=	2
constant long UNDERLINE_SINGLEACC = 33//0x21 
constant long UNDERLINE_DOUBLEACC = 34//0x22};


CONSTANT LONG PAPER_DEFAULT = 0
CONSTANT LONG PAPER_LETTER = 1
CONSTANT LONG PAPER_LETTERSMALL = 2
CONSTANT LONG PAPER_TABLOID = 3
CONSTANT LONG PAPER_LEDGER = 4
CONSTANT LONG PAPER_LEGAL = 5
CONSTANT LONG PAPER_STATEMENT = 6
CONSTANT LONG PAPER_EXECUTIVE = 7
CONSTANT LONG PAPER_A3 = 8
CONSTANT LONG PAPER_A4 = 9
CONSTANT LONG PAPER_A4SMALL = 10
CONSTANT LONG PAPER_A5 = 11
CONSTANT LONG PAPER_B4 = 12
CONSTANT LONG PAPER_B5 = 13
CONSTANT LONG PAPER_FOLIO = 14
CONSTANT LONG PAPER_QUATRO = 15
CONSTANT LONG PAPER_10x14 = 16
CONSTANT LONG PAPER_10x17 = 17
CONSTANT LONG PAPER_NOTE = 18
CONSTANT LONG PAPER_ENVELOPE_9 = 19
CONSTANT LONG PAPER_ENVELOPE_10 = 20
CONSTANT LONG PAPER_ENVELOPE_11 = 21
CONSTANT LONG PAPER_ENVELOPE_12 = 22
CONSTANT LONG PAPER_ENVELOPE_14 = 23
CONSTANT LONG PAPER_C_SIZE = 24
CONSTANT LONG PAPER_D_SIZE = 25
CONSTANT LONG PAPER_E_SIZE = 26
CONSTANT LONG PAPER_ENVELOPE_DL = 27
CONSTANT LONG PAPER_ENVELOPE_C5 = 28
CONSTANT LONG PAPER_ENVELOPE_C3 = 29
CONSTANT LONG PAPER_ENVELOPE_C4 = 30
CONSTANT LONG PAPER_ENVELOPE_C6 = 31
CONSTANT LONG PAPER_ENVELOPE_C65 = 32
CONSTANT LONG PAPER_ENVELOPE_B4 = 33
CONSTANT LONG PAPER_ENVELOPE_B5 = 34
CONSTANT LONG PAPER_ENVELOPE_B6 = 35
CONSTANT LONG PAPER_ENVELOPE = 36
CONSTANT LONG PAPER_ENVELOPE_MONARCH = 37
CONSTANT LONG PAPER_US_ENVELOPE = 38
CONSTANT LONG PAPER_FANFOLD = 39
CONSTANT LONG PAPER_GERMAN_STD_FANFOLD = 40
CONSTANT LONG PAPER_GERMAN_LEGAL_FANFOLD = 41
CONSTANT LONG PAPER_B4_ISO = 42
CONSTANT LONG PAPER_JAPANESE_POSTCARD = 43
CONSTANT LONG PAPER_9x11 = 44
CONSTANT LONG PAPER_10x11 = 45
CONSTANT LONG PAPER_15x11 = 46
CONSTANT LONG PAPER_ENVELOPE_INVITE = 47
CONSTANT LONG PAPER_US_LETTER_EXTRA  = 50
CONSTANT LONG PAPER_US_LEGAL_EXTRA = 51
CONSTANT LONG PAPER_US_TABLOID_EXTRA = 52
CONSTANT LONG PAPER_A4_EXTRA = 53
CONSTANT LONG PAPER_LETTER_TRANSVERSE = 54
CONSTANT LONG PAPER_A4_TRANSVERSE = 55
CONSTANT LONG PAPER_LETTER_EXTRA_TRANSVERSE = 56
CONSTANT LONG PAPER_SUPERA = 57
CONSTANT LONG PAPER_SUPERB = 58
CONSTANT LONG PAPER_US_LETTER_PLUS = 59
CONSTANT LONG PAPER_A4_PLUS = 60
CONSTANT LONG PAPER_A5_TRANSVERSE = 61
CONSTANT LONG PAPER_B5_TRANSVERSE = 62
CONSTANT LONG PAPER_A3_EXTRA = 63
CONSTANT LONG PAPER_A5_EXTRA = 64
CONSTANT LONG PAPER_B5_EXTRA = 65
CONSTANT LONG PAPER_A2 = 66
CONSTANT LONG PAPER_A3_TRANSVERSE = 67
CONSTANT LONG PAPER_A3_EXTRA_TRANSVERSE = 68
CONSTANT LONG PAPER_JAPANESE_DOUBLE_POSTCARD = 69
CONSTANT LONG PAPER_A6 = 70
CONSTANT LONG PAPER_JAPANESE_ENVELOPE_KAKU2 = 71
CONSTANT LONG PAPER_JAPANESE_ENVELOPE_KAKU3 = 72
CONSTANT LONG PAPER_JAPANESE_ENVELOPE_CHOU3 = 73
CONSTANT LONG PAPER_JAPANESE_ENVELOPE_CHOU4 = 74
CONSTANT LONG PAPER_LETTER_ROTATED = 75
CONSTANT LONG PAPER_A3_ROTATED = 76
CONSTANT LONG PAPER_A4_ROTATED = 77
CONSTANT LONG PAPER_A5_ROTATED = 78
CONSTANT LONG PAPER_B4_ROTATED = 79
CONSTANT LONG PAPER_B5_ROTATED = 80
CONSTANT LONG PAPER_JAPANESE_POSTCARD_ROTATED = 81
CONSTANT LONG PAPER_DOUBLE_JAPANESE_POSTCARD_ROTATED = 82
CONSTANT LONG PAPER_A6_ROTATED = 83
CONSTANT LONG PAPER_JAPANESE_ENVELOPE_KAKU2_ROTATED = 84
CONSTANT LONG PAPER_JAPANESE_ENVELOPE_KAKU3_ROTATED = 85
CONSTANT LONG PAPER_JAPANESE_ENVELOPE_CHOU3_ROTATED = 86
CONSTANT LONG PAPER_JAPANESE_ENVELOPE_CHOU4_ROTATED = 87
CONSTANT LONG PAPER_B6 = 88
CONSTANT LONG PAPER_B6_ROTATED = 89
CONSTANT LONG PAPER_12x11 = 90
CONSTANT LONG PAPER_JAPANESE_ENVELOPE_YOU4 = 91
CONSTANT LONG PAPER_JAPANESE_ENVELOPE_YOU4_ROTATED = 92
CONSTANT LONG PAPER_PRC16K = 93
CONSTANT LONG PAPER_PRC32K = 94
CONSTANT LONG PAPER_PRC32K_BIG = 95
CONSTANT LONG PAPER_PRC_ENVELOPE1 = 96
CONSTANT LONG PAPER_PRC_ENVELOPE2 = 97
CONSTANT LONG PAPER_PRC_ENVELOPE3 = 98
CONSTANT LONG PAPER_PRC_ENVELOPE4 = 99
CONSTANT LONG PAPER_PRC_ENVELOPE5 = 100
CONSTANT LONG PAPER_PRC_ENVELOPE6 = 101
CONSTANT LONG PAPER_PRC_ENVELOPE7 = 102
CONSTANT LONG PAPER_PRC_ENVELOPE8 = 103
CONSTANT LONG PAPER_PRC_ENVELOPE9 = 104
CONSTANT LONG PAPER_PRC_ENVELOPE10 = 105
CONSTANT LONG PAPER_PRC16K_ROTATED = 106
CONSTANT LONG PAPER_PRC32K_ROTATED = 107
CONSTANT LONG PAPER_PRC32KBIG_ROTATED = 108
CONSTANT LONG PAPER_PRC_ENVELOPE1_ROTATED = 109
CONSTANT LONG PAPER_PRC_ENVELOPE2_ROTATED = 110
CONSTANT LONG PAPER_PRC_ENVELOPE3_ROTATED = 111
CONSTANT LONG PAPER_PRC_ENVELOPE4_ROTATED = 112
CONSTANT LONG PAPER_PRC_ENVELOPE5_ROTATED = 113
CONSTANT LONG PAPER_PRC_ENVELOPE6_ROTATED = 114
CONSTANT LONG PAPER_PRC_ENVELOPE7_ROTATED = 115
CONSTANT LONG PAPER_PRC_ENVELOPE8_ROTATED = 116
CONSTANT LONG PAPER_PRC_ENVELOPE9_ROTATED = 117
CONSTANT LONG PAPER_PRC_ENVELOPE10_ROTATED = 118


CONSTANT LONG FILLPATTERN_NONE				=	0
CONSTANT LONG FILLPATTERN_SOLID				=	1
CONSTANT LONG FILLPATTERN_GRAY50				= 	2
CONSTANT LONG FILLPATTERN_GRAY75				=	3
CONSTANT LONG FILLPATTERN_GRAY25				=	4
CONSTANT LONG FILLPATTERN_HORSTRIPE			=	5
CONSTANT LONG FILLPATTERN_VERSTRIPE			=	6
CONSTANT LONG FILLPATTERN_REVDIAGSTRIPE	=	7
CONSTANT LONG FILLPATTERN_DIAGSTRIPE		=	8
CONSTANT LONG FILLPATTERN_DIAGCROSSHATCH	=	9	
CONSTANT LONG FILLPATTERN_THICKDIAGCROSSHATCH	=	10 
CONSTANT LONG FILLPATTERN_THINHORSTRIPE			=	11
CONSTANT LONG FILLPATTERN_THINVERSTRIPE			=	12
CONSTANT LONG FILLPATTERN_THINREVDIAGSTRIPE		=	13
CONSTANT LONG FILLPATTERN_THINDIAGSTRIPE			=	14
CONSTANT LONG FILLPATTERN_THINHORCROSSHATCH		=	15
CONSTANT LONG FILLPATTERN_THINDIAGCROSSHATCH		=	16
CONSTANT LONG FILLPATTERN_GRAY12P5					=	17
CONSTANT LONG FILLPATTERN_GRAY6P25					=	18


CONSTANT LONG PICTURETYPE_PNG = 0
CONSTANT LONG PICTURETYPE_JPEG = 1
CONSTANT LONG PICTURETYPE_GIF = 2
CONSTANT LONG PICTURETYPE_WMF = 3
CONSTANT LONG PICTURETYPE_DIB = 4
CONSTANT LONG PICTURETYPE_EMF = 5
CONSTANT LONG PICTURETYPE_PICT = 6
CONSTANT LONG PICTURETYPE_TIFF = 7
CONSTANT LONG PICTURETYPE_ERROR = 255


CONSTANT LONG SHEETTYPE_SHEET = 0
CONSTANT LONG SHEETTYPE_CHART = 1
CONSTANT LONG SHEETTYPE_UNKNOWN = 2



CONSTANT LONG ERRORTYPE_NULL = 0
CONSTANT LONG ERRORTYPE_DIV_0 = 7
CONSTANT LONG ERRORTYPE_VALUE = 15
CONSTANT LONG ERRORTYPE_REF = 23
CONSTANT LONG ERRORTYPE_NAME = 29
CONSTANT LONG ERRORTYPE_NUM = 36
CONSTANT LONG ERRORTYPE_NA = 42
CONSTANT LONG ERRORTYPE_NOERROR = 255

CONSTANT LONG SHEETSTATE_VISIBLE = 0
CONSTANT LONG SHEETSTATE_HIDDEN = 1
CONSTANT LONG SHEETSTATE_VERYHIDDEN = 2



CONSTANT LONG SCOPE_UNDEFINED = -2
CONSTANT LONG SCOPE_WORKBOOK = -1



CONSTANT LONG POSITION_MOVE_AND_SIZE = 0
CONSTANT LONG POSITION_ONLY_MOVE = 1
CONSTANT LONG POSITION_ABSOLUTE = 2


CONSTANT LONG OPERATOR_EQUAL = 0
CONSTANT LONG OPERATOR_GREATER_THAN = 1
CONSTANT LONG OPERATOR_GREATER_THAN_OR_EQUAL = 2
CONSTANT LONG OPERATOR_LESS_THAN = 3
CONSTANT LONG OPERATOR_LESS_THAN_OR_EQUAL = 4
CONSTANT LONG OPERATOR_NOT_EQUAL = 5


CONSTANT LONG FILTER_VALUE = 0
CONSTANT LONG FILTER_TOP10 = 1
CONSTANT LONG FILTER_CUSTOM = 2
CONSTANT LONG FILTER_DYNAMIC = 3
CONSTANT LONG FILTER_COLOR = 4
CONSTANT LONG FILTER_ICON = 5
CONSTANT LONG FILTER_EXT = 6
CONSTANT LONG FILTER_NOT_SET = 7

CONSTANT LONG IERR_NO_ERROR = 0
CONSTANT LONG IERR_EVAL_ERROR = 1
CONSTANT LONG IERR_EMPTY_CELLREF = 2
CONSTANT LONG IERR_NUMBER_STORED_AS_TEXT = 4
CONSTANT LONG IERR_INCONSIST_RANGE = 8
CONSTANT LONG IERR_INCONSIST_FMLA = 16
CONSTANT LONG IERR_TWODIG_TEXTYEAR = 32
CONSTANT LONG IERR_UNLOCK_FMLA = 64
CONSTANT LONG IERR_DATA_VALIDATION = 128


CONSTANT LONG PROT_DEFAULT = -1
CONSTANT LONG PROT_ALL = 0
CONSTANT LONG PROT_OBJECTS = 1
CONSTANT LONG PROT_SCENARIOS = 2
CONSTANT LONG PROT_FORMAT_CELLS = 4
CONSTANT LONG PROT_FORMAT_COLUMNS = 8
CONSTANT LONG PROT_FORMAT_ROWS = 16
CONSTANT LONG PROT_INSERT_COLUMNS = 32
CONSTANT LONG PROT_INSERT_ROWS = 64
CONSTANT LONG PROT_INSERT_HYPERLINKS = 128
CONSTANT LONG PROT_DELETE_COLUMNS = 256
CONSTANT LONG PROT_DELETE_ROWS = 512
CONSTANT LONG PROT_SEL_LOCKED_CELLS = 1024
CONSTANT LONG PROT_SORT = 2048
CONSTANT LONG PROT_AUTOFILTER = 4096
CONSTANT LONG PROT_PIVOTTABLES = 8192
CONSTANT LONG PROT_SEL_UNLOCKED_CELLS = 16384

CONSTANT LONG VALIDATION_TYPE_NONE = 0
CONSTANT LONG VALIDATION_TYPE_WHOLE = 1
CONSTANT LONG VALIDATION_TYPE_DECIMAL = 2
CONSTANT LONG VALIDATION_TYPE_LIST = 3
CONSTANT LONG VALIDATION_TYPE_DATE = 4
CONSTANT LONG VALIDATION_TYPE_TIME = 5
CONSTANT LONG VALIDATION_TYPE_TEXTLENGTH = 6
CONSTANT LONG VALIDATION_TYPE_CUSTOM = 7


CONSTANT LONG VALIDATION_OP_BETWEEN = 0
CONSTANT LONG VALIDATION_OP_NOTBETWEEN = 1
CONSTANT LONG VALIDATION_OP_EQUAL = 2
CONSTANT LONG VALIDATION_OP_NOTEQUAL = 3
CONSTANT LONG VALIDATION_OP_LESSTHAN = 4
CONSTANT LONG VALIDATION_OP_LESSTHANOREQUAL = 5
CONSTANT LONG VALIDATION_OP_GREATERTHAN = 6
CONSTANT LONG VALIDATION_OP_GREATERTHANOREQUAL = 7

CONSTANT LONG VALIDATION_ERRSTYLE_STOP = 0
CONSTANT LONG VALIDATION_ERRSTYLE_WARNING = 1
CONSTANT LONG VALIDATION_ERRSTYLE_INFORMATION = 2


CONSTANT LONG CALCMODE_MANUAL = 0
CONSTANT LONG CALCMODE_AUTO = 1
CONSTANT LONG CALCMODE_AUTONOTABLE = 2

CONSTANT LONG CHECKEDTYPE_UNCHECKED = 0
CONSTANT LONG CHECKEDTYPE_CHECKED = 1
CONSTANT LONG CHECKEDTYPE_MIXED = 2


CONSTANT LONG OBJECT_UNKNOWN = 0
CONSTANT LONG OBJECT_BUTTON = 1
CONSTANT LONG OBJECT_CHECKBOX = 2
CONSTANT LONG OBJECT_DROP = 3
CONSTANT LONG OBJECT_GBOX = 4
CONSTANT LONG OBJECT_LABEL = 5
CONSTANT LONG OBJECT_LIST = 6
CONSTANT LONG OBJECT_RADIO = 7
CONSTANT LONG OBJECT_SCROLL = 8
CONSTANT LONG OBJECT_SPIN = 9
CONSTANT LONG OBJECT_EDITBOX = 10
CONSTANT LONG OBJECT_DIALOG = 11

CONSTANT LONG SCRIPT_NORMAL = 0
CONSTANT LONG SCRIPT_SUPER = 1
CONSTANT LONG SCRIPT_SUB = 2

CONSTANT LONG CFORMAT_BEGINWITH = 0
CONSTANT LONG CFORMAT_CONTAINSBLANKS = 1
CONSTANT LONG CFORMAT_CONTAINSERRORS = 2
CONSTANT LONG CFORMAT_CONTAINSTEXT = 3
CONSTANT LONG CFORMAT_DUPLICATEVALUES = 4
CONSTANT LONG CFORMAT_ENDSWITH = 5
CONSTANT LONG CFORMAT_EXPRESSION = 6
CONSTANT LONG CFORMAT_NOTCONTAINSBLANKS = 7
CONSTANT LONG CFORMAT_NOTCONTAINSERRORS = 8
CONSTANT LONG CFORMAT_NOTCONTAINSTEXT = 9
CONSTANT LONG CFORMAT_UNIQUEVALUES = 10



CONSTANT LONG CFOPERATOR_LESSTHAN = 0
CONSTANT LONG CFOPERATOR_LESSTHANOREQUAL = 1
CONSTANT LONG CFOPERATOR_EQUAL = 2
CONSTANT LONG CFOPERATOR_NOTEQUAL = 3
CONSTANT LONG CFOPERATOR_GREATERTHANOREQUAL = 4
CONSTANT LONG CFOPERATOR_GREATERTHAN = 5
CONSTANT LONG CFOPERATOR_BETWEEN = 6
CONSTANT LONG CFOPERATOR_NOTBETWEEN = 7
CONSTANT LONG CFOPERATOR_CONTAINSTEXT = 8
CONSTANT LONG CFOPERATOR_NOTCONTAINS = 9
CONSTANT LONG CFOPERATOR_BEGINSWITH = 10
CONSTANT LONG CFOPERATOR_ENDSWITH = 11



CONSTANT LONG CFTP_LAST7DAYS = 0
CONSTANT LONG CFTP_LASTMONTH = 1
CONSTANT LONG CFTP_LASTWEEK = 2
CONSTANT LONG CFTP_NEXTMONTH = 3
CONSTANT LONG CFTP_NEXTWEEK = 4
CONSTANT LONG CFTP_THISMONTH = 5
CONSTANT LONG CFTP_THISWEEK = 6
CONSTANT LONG CFTP_TODAY = 7
CONSTANT LONG CFTP_TOMORROW = 8
CONSTANT LONG CFTP_YESTERDAY = 9


CONSTANT LONG CFVO_MIN = 0
CONSTANT LONG CFVO_MAX = 1
CONSTANT LONG CFVO_FORMULA = 2
CONSTANT LONG CFVO_NUMBER = 3
CONSTANT LONG CFVO_PERCENT = 4
CONSTANT LONG CFVO_PERCENTILE = 5


end variables
on cxls.create
call super::create
TriggerEvent( this, "constructor" )
end on

on cxls.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on

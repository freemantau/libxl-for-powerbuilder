HA$PBExportHeader$nx_xlbook.sru
forward
global type nx_xlbook from nonvisualobject
end type
end forward

global type nx_xlbook from nonvisualobject native "tplib_rs.dll"
public function  boolean load(readonly string filename)
public function  boolean save(readonly string filename)
public function  boolean save(readonly string filename,readonly boolean usetempfile)
public function  nx_xlsheet getsheet(readonly long index)
public function  nx_xlsheet addsheet(readonly string sheetname)
public function  boolean loadusingtempfile(readonly string filename,readonly string tempfile)
public function  boolean saveusingtempfile(readonly string filename,readonly boolean usetempfile)
public function  boolean loadpartially(readonly string filename,readonly long sheetindex,readonly long firstrow,readonly long lastrow)
public function  boolean loadpartiallyusingtempfile(readonly string filename,readonly long sheetindex,readonly long firstrow,readonly long lastrow,readonly string tempfile)
public function  long loadwithoutemptycells(readonly string filename)
public function  boolean loadinfo(readonly string filename)
public function  boolean loadraw(readonly blob data,readonly ulong size)
public function  boolean loadrawpartially(readonly blob data,readonly ulong size,readonly long sheetindex,readonly long firstrow,readonly long lastrow)
public function  boolean saveraw(ref blob data,ref ulong size)
public function  nx_xlsheet insertsheet(readonly long index,readonly string name,nx_xlsheet sheet)
public function  string getsheetname(readonly long index)
public function  boolean movesheet(readonly long srcindex,readonly long dstindex)
public function  boolean delsheet(readonly long index)
public function  long getsheetcount()
public function  nx_xlformat addformat()
public function  nx_xlformat addformat(nx_xlformat format)
public function  nx_xlfont addfont()
public function  nx_xlfont addfont(nx_xlfont font)
public function  nx_xlrichstring addrichstring()
public function  long addcustomnumformat(readonly string customNumFormat)
public function  string customnumformat(readonly long fmt)
public function  nx_xlformat format(readonly long index)
public function  long formatsize()
public function  nx_xlfont font(readonly long index)
public function  long fontsize()
public function  nx_xlconditionalformat addconditionalformat()
public function  double datepack(readonly long yy,readonly long mon,readonly long dd,readonly long hh,readonly long min,readonly long sec,readonly long msec)
public function  boolean dateunpack(readonly double value,ref long yy,ref long mon,ref long dd,ref long hh,ref long min,ref long sec,ref long msec)
public function  long colorpack(readonly long red,readonly long green,readonly long blue)
subroutine colorunpack(readonly long color,ref long red,ref long green,ref long blue)
public function  long activesheet()
subroutine setactivesheet(readonly long index)
public function  long picturesize()
public function  long getpicture(readonly long index,ref blob data,ref ulong size)
public function  long addpicture(readonly string filename)
public function  long addpicture2(readonly blob data,readonly ulong size)
public function  long addpictureaslink(readonly string filename)
public function  long addpictureaslink(readonly string filename,readonly boolean isinsert)
public function string defaultfont(ref long fontsize)
subroutine setdefaultfont(readonly string fontname,readonly long nfontsize)
public function  boolean refr1c1()
subroutine setrefr1c1(readonly boolean refR1C1)
public function  boolean rgbmode()
subroutine setrgbmode()
subroutine setrgbmode(readonly boolean brgbmode)
public function  boolean calcmode()
subroutine setcalcmode(readonly boolean bcalcmode)
public function  long version()
public function  long biffversionversion()
public function  boolean isdate1904()
subroutine setdate1904(readonly boolean bdate1904)
public function  boolean istemplate()
subroutine settemplate(readonly boolean btemplate)
public function  boolean iswriteprotected()
subroutine setkey(readonly string name,readonly string key)
public function  long setlocale(readonly blob locale)
public function string errormessage()
subroutine close()
public function boolean createxls(int xltype)
public function long getrowcount(readonly long sheetindex)
public function long getcolcount(readonly long sheetindex)
public function long getrow(readonly long sheetindex,readonly long row,ref string data[])
public function boolean isempty()
end type
global nx_xlbook nx_xlbook

type variables
//
end variables

on nx_xlbook.create
call super::create
TriggerEvent( this, "constructor" )
end on

on nx_xlbook.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on


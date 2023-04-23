HA$PBExportHeader$nx_xlformat.sru
forward
global type nx_xlformat from nonvisualobject
end type
end forward

global type nx_xlformat from nonvisualobject native "tplib_rs.dll"
public function nx_xlfont font()
public function boolean setfont(nx_xlfont font)
public function long numformat()
subroutine setnumformat(readonly long nval)
public function long alignh()
subroutine setalignh(readonly long nval)
public function long alignv()
subroutine setalignv(readonly long nval)
public function boolean wrap()
subroutine setwrap(readonly boolean nval)
public function long rotation()
public function long setrotation(readonly long nval)
public function long indent()
subroutine setindent(readonly long nval)
public function boolean shrinktofix()
subroutine setshrinktofix(readonly boolean nval)
subroutine setborder(readonly long nval)
subroutine setbordercolor(readonly long nval)
public function long borderleft()
subroutine setborderleft(readonly long nval)
public function long setborderleft()
subroutine setborderright(readonly long nval)
public function long bordertop()
subroutine setbordertop(readonly long nval)
public function long borderbottom()
subroutine setborderbottom(readonly long nval)
public function long borderleftcolor()
subroutine setborderleftcolor(readonly long nval)
public function long borderrightcolor()
subroutine setborderrightcolor(readonly long nval)
public function long bordertopcolor()
subroutine setbordertopcolor(readonly long nval)
public function long borderbottomcolor()
subroutine setborderbottomcolor(readonly long nval)
public function long borderdiagonal()
subroutine setborderdiagonal(readonly long nval)
public function long borderdiagonalstyle()
subroutine setborderdiagonalstyle(readonly long nval)
public function long borderdiagonalcolor()
subroutine setborderdiagonalcolor(readonly long nval)
public function long fillpattern()
subroutine setfillpattern(readonly long nval)
public function long patternforegroundcolor()
subroutine setpatternforegroundcolor(readonly long nval)
public function long patternbackgroundcolor()
subroutine setpatternbackgroundcolor(readonly long nval)
public function boolean locked()
subroutine setlocked(readonly boolean nval)
public function boolean hidden()
subroutine sethidden(readonly boolean nval)
public function boolean isempty()
end type
global nx_xlformat nx_xlformat

on nx_xlformat.create
call super::create
TriggerEvent( this, "constructor" )
end on

on nx_xlformat.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on


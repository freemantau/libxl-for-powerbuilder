HA$PBExportHeader$nx_xlconditionalformat.sru
forward
global type nx_xlconditionalformat from nonvisualobject
end type
end forward

global type nx_xlconditionalformat from nonvisualobject native "tplib_rs.dll"
public function nx_xlfont font()
public function long numformat()
subroutine setnumformat(readonly long numformat)
public function string customnumformat()
subroutine setcustomnumformat(readonly string customnumformat)
subroutine setborder()
subroutine setborder(readonly long style)
subroutine setbordercolor(readonly long color)
public function long borderleft()
subroutine setborderleft()
subroutine setborderleft(readonly long style)
public function long borderright()
subroutine setborderright()
subroutine setborderright(readonly long style)
public function long bordertop()
subroutine setbordertop()
subroutine setbordertop(readonly long style)
public function long borderbottom()
subroutine setborderbottom()
subroutine setborderbottom(readonly long style)
public function long borderleftcolor()
subroutine setborderleftcolor(readonly long color)
public function long borderrightcolor()
subroutine setborderrightcolor(readonly long color)
public function long bordertopcolor()
subroutine setbordertopcolor(readonly long color)
public function long borderbottomcolor()
subroutine setborderbottomcolor(readonly long color)
public function long fillpattern()
subroutine setfillpattern(readonly long pattern)
public function long patternforegroundcolor()
subroutine setpatternforegroundcolor(readonly long color)
public function long patternbackgroundcolor()
subroutine setpatternbackgroundcolor(readonly long color)
public function boolean isempty()
end type
global nx_xlconditionalformat nx_xlconditionalformat

on nx_xlconditionalformat.create
call super::create
TriggerEvent( this, "constructor" )
end on

on nx_xlconditionalformat.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on


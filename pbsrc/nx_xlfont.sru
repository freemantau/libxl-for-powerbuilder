HA$PBExportHeader$nx_xlfont.sru
forward
global type nx_xlfont from nonvisualobject
end type
end forward

global type nx_xlfont from nonvisualobject native "tplib_rs.dll"
public function long size()
subroutine setsize(readonly long size)
public function boolean italic()
subroutine setitalic(readonly boolean itealic)
public function boolean strikeout()
subroutine setstrikeout(readonly boolean bstrikeout)
public function long color()
subroutine setcolor(readonly long ncolor)
public function boolean bold()
subroutine setbold(readonly boolean bbold)
public function long script()
subroutine setscript(readonly long nscript)
public function long underline()
subroutine setunderline(readonly long bunderline)
public function string name()
subroutine setname(readonly string name)
public function boolean isempty()
end type
global nx_xlfont nx_xlfont

on nx_xlfont.create
call super::create
TriggerEvent( this, "constructor" )
end on

on nx_xlfont.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on


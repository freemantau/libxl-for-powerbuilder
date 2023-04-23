HA$PBExportHeader$nx_xlrichstring.sru
forward
global type nx_xlrichstring from nonvisualobject
end type
end forward

global type nx_xlrichstring from nonvisualobject native "tplib_rs.dll"
public function nx_xlfont addfont()
public function nx_xlfont addfont(nx_xlfont fnt)
subroutine addtext(readonly string text,nx_xlfont fnt)
public function string gettext(readonly long index,nx_xlfont fnt)
public function long textsize()
public function boolean isempty()
end type
global nx_xlrichstring nx_xlrichstring

on nx_xlrichstring.create
call super::create
TriggerEvent( this, "constructor" )
end on

on nx_xlrichstring.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on


HA$PBExportHeader$nx_xlformcontrol.sru
forward
global type nx_xlformcontrol from nonvisualobject
end type
end forward

global type nx_xlformcontrol from nonvisualobject native "tplib_rs.dll"
public function boolean isempty()
end type
global nx_xlformcontrol nx_xlformcontrol

on nx_xlformcontrol.create
call super::create
TriggerEvent( this, "constructor" )
end on

on nx_xlformcontrol.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on


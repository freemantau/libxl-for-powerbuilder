HA$PBExportHeader$nx_xlautofilter.sru
forward
global type nx_xlautofilter from nonvisualobject
end type
end forward

global type nx_xlautofilter from nonvisualobject native "tplib_rs.dll"
public function boolean getref(ref long rowfirst, ref long rowlast, ref long colfirst, ref long collast)
subroutine  setref(readonly long rowfirst, readonly long rowlast, readonly long colfirst, readonly long collast)
public function  nx_xlfiltercolumn  column(readonly long colid)
public function long columnsize()
public function  nx_xlfiltercolumn  columnbyindex(readonly long index)
public function boolean getsortrange(ref long rowfirst, ref long rowlast, ref long colfirst, ref long collast)
public function boolean getsort(ref long columnindex, ref boolean descending)
public function boolean setsort(readonly long columnindex)
public function boolean setsort(readonly long columnindex, readonly boolean descending)
public function boolean addsort(readonly long columnindex)
public function boolean addsort(readonly long columnindex, readonly boolean descending)
public function boolean isempty()
end type
global nx_xlautofilter nx_xlautofilter

on nx_xlautofilter.create
call super::create
TriggerEvent( this, "constructor" )
end on

on nx_xlautofilter.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on


HA$PBExportHeader$nx_xlfiltercolumn.sru
forward
global type nx_xlfiltercolumn from nonvisualobject
end type
end forward

global type nx_xlfiltercolumn from nonvisualobject native "tplib_rs.dll"
public function long Index()
public function long FilterType()
public function long FilterSize()
public function string Filter(readonly long index)
subroutine AddFilter(readonly string  value)
public function boolean GetTop10(ref double value, ref long top, ref long percent)
subroutine SetTop10(readonly double value)
subroutine SetTop10(readonly double value,readonly  boolean top)
subroutine SetTop10(readonly double value,readonly  boolean top,readonly  boolean percent)
public function boolean GetCustomFilter(ref long op1, ref string  v1, ref long op2, ref string  v2, ref long andOp)
subroutine SetCustomFilterEx(readonly long op1, readonly string  v1)
subroutine SetCustomFilterEx(readonly long op1, readonly string  v1,readonly  long op2)
subroutine SetCustomFilterEx(readonly long op1, readonly string  v1,readonly  long op2, readonly string  v2)
subroutine SetCustomFilterEx(readonly long op1, readonly string  v1,readonly  long op2, readonly string  v2,readonly  boolean andOp)
subroutine Clear()
public function boolean isempty()
end type
global nx_xlfiltercolumn nx_xlfiltercolumn

on nx_xlfiltercolumn.create
call super::create
TriggerEvent( this, "constructor" )
end on

on nx_xlfiltercolumn.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on


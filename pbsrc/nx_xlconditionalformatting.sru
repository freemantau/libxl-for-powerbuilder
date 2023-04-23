HA$PBExportHeader$nx_xlconditionalformatting.sru
forward
global type nx_xlconditionalformatting from nonvisualobject
end type
end forward

global type nx_xlconditionalformatting from nonvisualobject native "tplib_rs.dll"
subroutine addrange(readonly long rowfirst, readonly long rowlast, readonly long colfirst, readonly long collast)
subroutine addrule(readonly long type_, nx_xlconditionalformat cformat)
subroutine addrule(readonly long type_, nx_xlconditionalformat cformat, readonly string value)
subroutine addrule(readonly long type_, nx_xlconditionalformat cformat, readonly string value, readonly boolean stopiftrue)
subroutine addtoprule(nx_xlconditionalformat cformat, readonly long value)
subroutine addtoprule(nx_xlconditionalformat cformat, readonly long value, readonly long bottom)
subroutine addtoprule(nx_xlconditionalformat cformat, readonly long value, readonly long bottom, readonly long percent)
subroutine addtoprule(nx_xlconditionalformat cformat, readonly long value, readonly long bottom, readonly long percent, readonly boolean stopiftrue)
subroutine addopnumrule(readonly long op, nx_xlconditionalformat cformat, double value1)
subroutine addopnumrule(readonly long op, nx_xlconditionalformat cformat, double value1, double value2)
subroutine addopnumrule(readonly long op, nx_xlconditionalformat cformat, double value1, double value2, readonly boolean stopiftrue)
subroutine addopstrrule(readonly long op, nx_xlconditionalformat cformat, readonly string value1)
subroutine addopstrrule(readonly long op, nx_xlconditionalformat cformat, readonly string value1, readonly string value2)
subroutine addopstrrule(readonly long op, nx_xlconditionalformat cformat, readonly string value1, readonly string value2, readonly boolean stopiftrue)
subroutine addaboveaveragerule(nx_xlconditionalformat cformat)
subroutine addaboveaveragerule(nx_xlconditionalformat cformat, readonly long aboveaverage)
subroutine addaboveaveragerule(nx_xlconditionalformat cformat, readonly long aboveaverage, readonly long equalaverage)
subroutine addaboveaveragerule(nx_xlconditionalformat cformat, readonly long aboveaverage, readonly long equalaverage, readonly long stddev)
subroutine addaboveaveragerule(nx_xlconditionalformat cformat, readonly long aboveaverage, readonly long equalaverage, readonly long stddev, readonly boolean stopiftrue)
subroutine addtimeperiodrule(nx_xlconditionalformat cformat, readonly long timeperiod)
subroutine addtimeperiodrule(nx_xlconditionalformat cformat, readonly long timeperiod, readonly boolean stopiftrue)
subroutine add2colorscalerule(readonly long mincolor, readonly long maxcolor)
subroutine add2colorscalerule(readonly long mincolor, readonly long maxcolor, readonly long mintype)
subroutine add2colorscalerule(readonly long mincolor, readonly long maxcolor, readonly long mintype, double minvalue)
subroutine add2colorscalerule(readonly long mincolor, readonly long maxcolor, readonly long mintype, double minvalue, readonly long maxtype)
subroutine add2colorscalerule(readonly long mincolor, readonly long maxcolor, readonly long mintype, double minvalue, readonly long maxtype, double maxvalue)
subroutine add2colorscalerule(readonly long mincolor, readonly long maxcolor, readonly long mintype, double minvalue, readonly long maxtype, double maxvalue, readonly boolean stopiftrue)
subroutine add2colorscaleformularule(readonly long mincolor, readonly long maxcolor)
subroutine add2colorscaleformularule(readonly long mincolor, readonly long maxcolor, readonly long mintype)
subroutine add2colorscaleformularule(readonly long mincolor, readonly long maxcolor, readonly long mintype, readonly string minvalue)
subroutine add2colorscaleformularule(readonly long mincolor, readonly long maxcolor, readonly long mintype, readonly string minvalue, readonly long maxtype)
subroutine add2colorscaleformularule(readonly long mincolor, readonly long maxcolor, readonly long mintype, readonly string minvalue, readonly long maxtype, readonly string maxvalue)
subroutine add2colorscaleformularule(readonly long mincolor, readonly long maxcolor, readonly long mintype, readonly string minvalue, readonly long maxtype, readonly string maxvalue, readonly boolean stopiftrue)
subroutine add3colorscalerule(readonly long mincolor, readonly long midcolor, readonly long maxcolor)
subroutine add3colorscalerule(readonly long mincolor, readonly long midcolor, readonly long maxcolor, readonly long mintype)
subroutine add3colorscalerule(readonly long mincolor, readonly long midcolor, readonly long maxcolor, readonly long mintype, double minvalue)
subroutine add3colorscalerule(readonly long mincolor, readonly long midcolor, readonly long maxcolor, readonly long mintype, double minvalue, readonly long midtype)
subroutine add3colorscalerule(readonly long mincolor, readonly long midcolor, readonly long maxcolor, readonly long mintype, double minvalue, readonly long midtype, double midvalue)
subroutine add3colorscalerule(readonly long mincolor, readonly long midcolor, readonly long maxcolor, readonly long mintype, double minvalue, readonly long midtype, double midvalue, readonly long maxtype)
subroutine add3colorscalerule(readonly long mincolor, readonly long midcolor, readonly long maxcolor, readonly long mintype, double minvalue, readonly long midtype, double midvalue, readonly long maxtype, double maxvalue)
subroutine add3colorscalerule(readonly long mincolor, readonly long midcolor, readonly long maxcolor, readonly long mintype, double minvalue, readonly long midtype, double midvalue, readonly long maxtype, double maxvalue, readonly boolean stopiftrue)
subroutine add3colorscaleformularule(readonly long mincolor, readonly long midcolor, readonly long maxcolor)
subroutine add3colorscaleformularule(readonly long mincolor, readonly long midcolor, readonly long maxcolor, readonly long mintype)
subroutine add3colorscaleformularule(readonly long mincolor, readonly long midcolor, readonly long maxcolor, readonly long mintype, readonly string minvalue)
subroutine add3colorscaleformularule(readonly long mincolor, readonly long midcolor, readonly long maxcolor, readonly long mintype, readonly string minvalue, readonly long midtype)
subroutine add3colorscaleformularule(readonly long mincolor, readonly long midcolor, readonly long maxcolor, readonly long mintype, readonly string minvalue, readonly long midtype, readonly string midvalue)
subroutine add3colorscaleformularule(readonly long mincolor, readonly long midcolor, readonly long maxcolor, readonly long mintype, readonly string minvalue, readonly long midtype, readonly string midvalue, readonly long maxtype)
subroutine add3colorscaleformularule(readonly long mincolor, readonly long midcolor, readonly long maxcolor, readonly long mintype, readonly string minvalue, readonly long midtype, readonly string midvalue, readonly long maxtype, readonly string maxvalue)
subroutine add3colorscaleformularule(readonly long mincolor, readonly long midcolor, readonly long maxcolor, readonly long mintype, readonly string minvalue, readonly long midtype, readonly string midvalue, readonly long maxtype, readonly string maxvalue, readonly boolean stopiftrue)
public function boolean isempty()
end type
global nx_xlconditionalformatting nx_xlconditionalformatting

on nx_xlconditionalformatting.create
call super::create
TriggerEvent( this, "constructor" )
end on

on nx_xlconditionalformatting.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on


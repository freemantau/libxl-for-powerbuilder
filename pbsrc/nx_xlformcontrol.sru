HA$PBExportHeader$nx_xlformcontrol.sru
forward
global type nx_xlformcontrol from nonvisualobject
end type
end forward

global type nx_xlformcontrol from nonvisualobject native "tplib_rs.dll"
public function long  ObjectType()
public function long  Checked()
subroutine SetChecked(readonly long checked)
public function string FmlaGroup()
subroutine SetFmlaGroup(readonly string group)
public function string FmlaLink()
subroutine SetFmlaLink(readonly string link)
public function string FmlaRange()
subroutine SetFmlaRange(readonly string range)
public function string FmlaTxbx()
subroutine SetFmlaTxbx(readonly string txbx)
public function string Name()
public function string LinkedCell()
public function string ListFillRange()
public function string Macro()
public function string AltText()
public function boolean  Locked()
public function boolean  DefaultSize()
public function boolean  Print()
public function boolean  Disabled()
public function string Item(readonly long index)
public function long  ItemSize()
subroutine AddItem(readonly string value)
subroutine InsertItem(readonly long index, readonly string value)
subroutine ClearItems()
public function long  DropLines()
subroutine SetDropLines(readonly long lines)
public function long  Dx()
subroutine SetDx(readonly long dx)
public function boolean  FirstButton()
subroutine SetFirstButton(readonly boolean firstButton)
public function boolean  Horiz()
subroutine SetHoriz(readonly boolean horiz)
public function long  Inc()
subroutine SetInc(readonly long inc)
public function long  GetMax()
subroutine SetMax(readonly long max)
public function long  GetMin()
subroutine SetMin(readonly long min)
public function string MultiSel()
subroutine SetMultiSel(readonly string value)
public function long  Sel()
subroutine SetSel(readonly long sel)
public function boolean  FromAnchor(ref long col, ref long colOff, ref long row, ref long rowOff)
public function boolean  ToAnchor(ref long col, ref long colOff, ref long row, ref long rowOff)
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


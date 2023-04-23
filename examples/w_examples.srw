$PBExportHeader$w_examples.srw
forward
global type w_examples from window
end type
type cb_12 from commandbutton within w_examples
end type
type cb_11 from commandbutton within w_examples
end type
type cb_10 from commandbutton within w_examples
end type
type cb_9 from commandbutton within w_examples
end type
type st_2 from statictext within w_examples
end type
type cb_8 from commandbutton within w_examples
end type
type cb_7 from commandbutton within w_examples
end type
type cb_6 from commandbutton within w_examples
end type
type cb_5 from commandbutton within w_examples
end type
type cb_4 from commandbutton within w_examples
end type
type cb_3 from commandbutton within w_examples
end type
type cb_2 from commandbutton within w_examples
end type
type cb_1 from commandbutton within w_examples
end type
type st_1 from statictext within w_examples
end type
type cb_76 from commandbutton within w_examples
end type
type cb_75 from commandbutton within w_examples
end type
type cb_74 from commandbutton within w_examples
end type
type cb_73 from commandbutton within w_examples
end type
type cb_72 from commandbutton within w_examples
end type
type st_6 from statictext within w_examples
end type
type cb_71 from commandbutton within w_examples
end type
type cb_70 from commandbutton within w_examples
end type
type cb_69 from commandbutton within w_examples
end type
type cb_68 from commandbutton within w_examples
end type
type cb_67 from commandbutton within w_examples
end type
type cb_66 from commandbutton within w_examples
end type
type cb_65 from commandbutton within w_examples
end type
type cb_64 from commandbutton within w_examples
end type
type cb_63 from commandbutton within w_examples
end type
type cb_56 from commandbutton within w_examples
end type
type cb_55 from commandbutton within w_examples
end type
type st_5 from statictext within w_examples
end type
end forward

global type w_examples from window
integer width = 3954
integer height = 2476
boolean titlebar = true
string title = "tplib_rs"
boolean controlmenu = true
boolean minbox = true
boolean maxbox = true
boolean resizable = true
long backcolor = 67108864
string icon = "AppIcon!"
boolean center = true
cb_12 cb_12
cb_11 cb_11
cb_10 cb_10
cb_9 cb_9
st_2 st_2
cb_8 cb_8
cb_7 cb_7
cb_6 cb_6
cb_5 cb_5
cb_4 cb_4
cb_3 cb_3
cb_2 cb_2
cb_1 cb_1
st_1 st_1
cb_76 cb_76
cb_75 cb_75
cb_74 cb_74
cb_73 cb_73
cb_72 cb_72
st_6 st_6
cb_71 cb_71
cb_70 cb_70
cb_69 cb_69
cb_68 cb_68
cb_67 cb_67
cb_66 cb_66
cb_65 cb_65
cb_64 cb_64
cb_63 cb_63
cb_56 cb_56
cb_55 cb_55
st_5 st_5
end type
global w_examples w_examples

type variables
string is_path


protected:
constant string #NAME = "{name}"
constant string #KEY = "{key}"
end variables

on w_examples.create
this.cb_12=create cb_12
this.cb_11=create cb_11
this.cb_10=create cb_10
this.cb_9=create cb_9
this.st_2=create st_2
this.cb_8=create cb_8
this.cb_7=create cb_7
this.cb_6=create cb_6
this.cb_5=create cb_5
this.cb_4=create cb_4
this.cb_3=create cb_3
this.cb_2=create cb_2
this.cb_1=create cb_1
this.st_1=create st_1
this.cb_76=create cb_76
this.cb_75=create cb_75
this.cb_74=create cb_74
this.cb_73=create cb_73
this.cb_72=create cb_72
this.st_6=create st_6
this.cb_71=create cb_71
this.cb_70=create cb_70
this.cb_69=create cb_69
this.cb_68=create cb_68
this.cb_67=create cb_67
this.cb_66=create cb_66
this.cb_65=create cb_65
this.cb_64=create cb_64
this.cb_63=create cb_63
this.cb_56=create cb_56
this.cb_55=create cb_55
this.st_5=create st_5
this.Control[]={this.cb_12,&
this.cb_11,&
this.cb_10,&
this.cb_9,&
this.st_2,&
this.cb_8,&
this.cb_7,&
this.cb_6,&
this.cb_5,&
this.cb_4,&
this.cb_3,&
this.cb_2,&
this.cb_1,&
this.st_1,&
this.cb_76,&
this.cb_75,&
this.cb_74,&
this.cb_73,&
this.cb_72,&
this.st_6,&
this.cb_71,&
this.cb_70,&
this.cb_69,&
this.cb_68,&
this.cb_67,&
this.cb_66,&
this.cb_65,&
this.cb_64,&
this.cb_63,&
this.cb_56,&
this.cb_55,&
this.st_5}
end on

on w_examples.destroy
destroy(this.cb_12)
destroy(this.cb_11)
destroy(this.cb_10)
destroy(this.cb_9)
destroy(this.st_2)
destroy(this.cb_8)
destroy(this.cb_7)
destroy(this.cb_6)
destroy(this.cb_5)
destroy(this.cb_4)
destroy(this.cb_3)
destroy(this.cb_2)
destroy(this.cb_1)
destroy(this.st_1)
destroy(this.cb_76)
destroy(this.cb_75)
destroy(this.cb_74)
destroy(this.cb_73)
destroy(this.cb_72)
destroy(this.st_6)
destroy(this.cb_71)
destroy(this.cb_70)
destroy(this.cb_69)
destroy(this.cb_68)
destroy(this.cb_67)
destroy(this.cb_66)
destroy(this.cb_65)
destroy(this.cb_64)
destroy(this.cb_63)
destroy(this.cb_56)
destroy(this.cb_55)
destroy(this.st_5)
end on

event open;is_path  = GetCurrentDirectory()

end event

type cb_12 from commandbutton within w_examples
integer x = 2455
integer y = 1280
integer width = 1339
integer height = 100
integer taborder = 330
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Highlighting alternating rows"
end type

event clicked;    nx_xlbook book
	 book = create nx_xlbook
	 book.setkey( #NAME,#KEY )
	 nx_xlsheet sheet


    book.setRgbMode();
    
    sheet = book.addSheet("my");

    nx_xlConditionalFormat cFormat
	 cFormat = book.addConditionalFormat();
    cFormat.setFillPattern(cxls.FILLPATTERN_SOLID);
    cFormat.setPatternBackgroundColor(book.colorPack(240, 240, 240));

    nx_xlConditionalFormatting cFormatting 
	 cFormatting= sheet.addConditionalFormatting();
    cFormatting.addRange(4, 20, 1, 10);
    cFormatting.addRule(cxls.CFORMAT_EXPRESSION, cFormat, "=MOD(ROW(),2)=0");

	long row,col
	 for row = 4 to 20
		for col = 1 to 10
			sheet.writeNum(row, col, double(row) + double(col));
		next
	next
			
   
    book.save("alternatingrows.xlsx");

    destroy book;
end event

type cb_11 from commandbutton within w_examples
integer x = 2455
integer y = 1156
integer width = 1339
integer height = 100
integer taborder = 330
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Highlighting cells that more than the specified value"
end type

event clicked;nx_xlbook book
nx_xlsheet sheet
book = create nx_xlbook


book.setRgbMode(true);

sheet = book.addSheet("my");

sheet.writeStr(2, 1, "Country");
sheet.writeStr(2, 2, "Road injures");
sheet.writeStr(3, 1, "USA");
sheet.writeStr(4, 1, "UK");
sheet.writeStr(5, 1, "Switzerland");
sheet.writeStr(6, 1, "Spain");
sheet.writeStr(7, 1, "Italy");
sheet.writeStr(8, 1, "London");
sheet.writeStr(9, 1, "Greece");
sheet.writeStr(10, 1, "Japan");

sheet.writeNum(3, 2, 64);
sheet.writeNum(4, 2, 94);
sheet.writeNum(5, 2, 88);
sheet.writeNum(6, 2, 93);
sheet.writeNum(7, 2, 86);
sheet.writeNum(8, 2, 75);
sheet.writeNum(9, 2, 67);
sheet.writeNum(10, 2, 91);

nx_xlConditionalFormat cFormat
cFormat = book.addConditionalFormat();
cFormat.setFillPattern(cxls.FILLPATTERN_SOLID);
cFormat.setPatternBackgroundColor( book.colorPack(/*0xFF*/255, /*0xEF*/239, /*0x9C*/156));

nx_xlConditionalFormatting cf
cf = sheet.addConditionalFormatting();
cf.addRange(3, 10, 2, 2);
cf.addOpNumRule(cxls.CFOPERATOR_GREATERTHAN, cFormat, 90);
	 
book.save("out.xlsx");

destroy book;
end event

type cb_10 from commandbutton within w_examples
integer x = 2455
integer y = 1036
integer width = 1294
integer height = 100
integer taborder = 320
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Creating a gradated color scale on the cells"
end type

event clicked;nx_xlbook book
nx_xlsheet sheet
book = create nx_xlbook


book.setRgbMode(true);

sheet = book.addSheet("my");

sheet.writeStr(2, 1, "Country");
sheet.writeStr(2, 2, "Road injures");
sheet.writeStr(3, 1, "USA");
sheet.writeStr(4, 1, "UK");
sheet.writeStr(5, 1, "Switzerland");
sheet.writeStr(6, 1, "Spain");
sheet.writeStr(7, 1, "Italy");
sheet.writeStr(8, 1, "London");
sheet.writeStr(9, 1, "Greece");
sheet.writeStr(10, 1, "Japan");

sheet.writeNum(3, 2, 64);
sheet.writeNum(4, 2, 94);
sheet.writeNum(5, 2, 88);
sheet.writeNum(6, 2, 93);
sheet.writeNum(7, 2, 86);
sheet.writeNum(8, 2, 75);
sheet.writeNum(9, 2, 67);
sheet.writeNum(10, 2, 91);

nx_xlConditionalFormatting cf 
cf = sheet.addConditionalFormatting();
cf.addRange(3, 10, 2, 2);
cf.add2ColorScaleRule(book.colorPack(/*0xFF*/255, /*0x71*/113, /*0x28*/40), book.colorPack(/*0xFF*/255, /*0xEF*/239, /*0x9C*/156));

book.save("out.xlsx");

destroy book;
end event

type cb_9 from commandbutton within w_examples
integer x = 2455
integer y = 916
integer width = 1294
integer height = 100
integer taborder = 340
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Highlighting cells that begin with the given text"
end type

event clicked;nx_xlbook book
nx_xlsheet sheet
book = create nx_xlbook


sheet = book.addSheet("my");

sheet.writeStr(2, 1, "Poltava");
sheet.writeStr(3, 1, "Athens");
sheet.writeStr(4, 1, "Thessaloniki");
sheet.writeStr(5, 1, "Kalamata");
sheet.writeStr(6, 1, "Amsterdam");
sheet.writeStr(7, 1, "Alexandroupolis");
sheet.writeStr(8, 1, "Kyiv");
sheet.writeStr(9, 1, "London");
sheet.writeStr(10, 1, "New York");

nx_xlConditionalFormat cFormat
cFormat = book.addConditionalFormat();
cFormat.font().setBold(true);

nx_xlConditionalFormatting cf 
cf = sheet.addConditionalFormatting();
cf.addRange(2, 10, 1, 1);
cf.addRule(cxls.CFORMAT_BEGINWITH, cFormat, "a");

book.save("out.xlsx");

destroy book;
end event

type st_2 from statictext within w_examples
integer x = 2455
integer y = 816
integer width = 1413
integer height = 76
integer textsize = -9
integer weight = 700
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 16777215
long backcolor = 134217741
string text = "Conditional formatting"
boolean focusrectangle = false
end type

type cb_8 from commandbutton within w_examples
integer x = 59
integer y = 2248
integer width = 1262
integer height = 100
integer taborder = 370
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Writing rich strings with multiple fonts in one cell"
end type

event clicked;nx_xlbook book
book = create nx_xlbook

nx_xlsheet sheet
nx_xlrichstring rs
nx_xlfont font1,font2,font3,font4,font5

    sheet = book.addSheet("my");

    rs = book.addRichString();

    font1 = rs.addFont();    
    font1.setColor(cxls.COLOR_RED);
    font1.setSize(24);

    font2 = rs.addFont();
    font2.setSize(24);
        
    font3 = rs.addFont();
    font3.setColor(cxls.COLOR_BLUE);
    font3.setSize(24);

    font4 = rs.addFont();
    font4.setColor(cxls.COLOR_GREEN);
    font4.setSize(24);

    font5 = rs.addFont();
    font5.setScript(cxls.SCRIPT_SUPER);
    font5.setSize(24);

    rs.addText("E", font1);
    rs.addText("=", font2);
    rs.addText("m", font3);
    rs.addText("c", font4);
    rs.addText("2", font5);   

    sheet.writeRichStr(5, 3, rs);

    book.save("richString.xlsx");

    destroy book;
end event

type cb_7 from commandbutton within w_examples
integer x = 59
integer y = 2128
integer width = 859
integer height = 100
integer taborder = 360
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Replacing characters in all sheets"
end type

type cb_6 from commandbutton within w_examples
integer x = 59
integer y = 2008
integer width = 667
integer height = 100
integer taborder = 350
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Sheet protection"
end type

event clicked;nx_xlbook book
nx_xlsheet sheet
nx_xlformat format
book = create nx_xlbook

sheet = book.addSheet("my");

format = book.addFormat();
format.setLocked(false);

sheet.writeStr(5, 1, "this cell can be changed!", format);
sheet.writeNum(6, 1, 100, format);

sheet.setProtect(true);

book.save("out.xlsx");

destroy book;
end event

type cb_5 from commandbutton within w_examples
integer x = 59
integer y = 1888
integer width = 1189
integer height = 100
integer taborder = 340
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Copying cells with formats to other workbook"
end type

event clicked;nx_xlbook srcBook,dstBook
nx_xlsheet srcSheet,dstSheet
nx_xlformat fmtnull
fmtnull = create nx_xlformat
srcBook = create nx_xlbook
dstBook  = create nx_xlbook
dstBook.setkey(#NAME,#KEY)
srcBook.load("data.xlsx");

srcSheet = srcBook.getSheet(0);
dstSheet = dstBook.addSheet("my");

long col,row
// setting column widths
for col = srcSheet.firstCol() to srcSheet.lastCol()-1
dstSheet.setCol(col, col, srcSheet.colWidth(col), fmtnull, srcSheet.colHidden(col));
next


for row = srcSheet.firstRow() to srcSheet.lastRow()-1
// setting row heights
	dstSheet.setRow(row, srcSheet.rowHeight(row), fmtnull, srcSheet.rowHidden(row));
	for col = srcSheet.firstCol() to srcSheet.lastCol() -1	
			// copying merging blocks
			long rowFirst, rowLast, colFirst, colLast;
			if (srcSheet.getMerge(row, col, ref rowFirst, ref rowLast, ref colFirst, ref colLast)) then							 
				 dstSheet.setMerge(rowFirst, rowLast, colFirst, colLast);              
			end if			
			// copying formats
			nx_xlformat srcFormat,dstFormat		
			srcFormat = srcSheet.cellFormat(row, col);              
			dstFormat = dstBook.addFormat(srcFormat);			
			// copying cell's values
			long ct 
			ct = srcSheet.cellType(row, col);
			choose case ct			
				 case cxls.CELLTYPE_NUMBER				 
					  double dvalue
					  dvalue = srcSheet.readNum(row, col);                         
					  dstSheet.writeNum(row, col, dvalue, dstFormat);		 
				 case cxls.CELLTYPE_BOOLEAN				 
					  boolean bvalue
					  bvalue = srcSheet.readBool(row, col);
					  dstSheet.writeBool(row, col, bvalue, dstFormat);				 
				 case cxls.CELLTYPE_STRING				 
					  string s
					  s = srcSheet.readStr(row, col);                
					  dstSheet.writeStr(row, col, s, dstFormat);					 				 
				 case cxls.CELLTYPE_BLANK                  			
					  srcSheet.readBlank(row, col);
					  dstSheet.writeBlank(row, col, dstFormat);            
			end choose  
	next
next

dstBook.save("dst.xlsx");

destroy dstBook

destroy srcBook
end event

type cb_4 from commandbutton within w_examples
integer x = 59
integer y = 1768
integer width = 974
integer height = 100
integer taborder = 330
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Writing to memory buffer"
end type

event clicked;nx_xlbook book
book = create nx_xlbook

blob rawdata
ulong rawlen

book.load( "group.xlsx")

book.saveraw( ref rawdata, ref rawlen)


messagebox(string(len(rawdata)),string(rawlen))

destroy book
end event

type cb_3 from commandbutton within w_examples
integer x = 59
integer y = 1644
integer width = 974
integer height = 100
integer taborder = 320
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Reading from memory buffer"
end type

event clicked;nx_xlbook book
book = create nx_xlbook

int ifile
blob rawdata

ifile = fileopen("group.xlsx",streammode!)
filereadex(ifile,rawdata)
fileclose(ifile)

book.loadraw( rawdata,len(rawdata)) 
messagebox('',book.getsheet( 0).getname( ))

destroy book
end event

type cb_2 from commandbutton within w_examples
integer x = 59
integer y = 1060
integer width = 974
integer height = 100
integer taborder = 330
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Grouping rows and columns"
end type

event clicked;nx_xlbook book
book = create nx_xlbook
nx_xlsheet sheet

sheet = book.addSheet("Sheet1");
long row,col   
for row = 1 to 29    
  for col = 0 to 9
		sheet.writeNum(row, col, rand(10) );
	next
next
  
sheet.groupCols(0, 1);
sheet.groupCols(4, 7, false);    

sheet.groupRows(3, 6);
sheet.groupRows(10, 25, false);    
sheet.groupRows(14, 16, false);
sheet.groupRows(19, 21, false);

book.save("group.xlsx");

destroy book
end event

type cb_1 from commandbutton within w_examples
integer x = 59
integer y = 472
integer width = 974
integer height = 100
integer taborder = 290
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Extracting pictures"
end type

event clicked;nx_xlbook book
book = create nx_xlbook
book.load("in.xlsx");

long i,nc
nc = book.picturesize( )
messagebox("picture count",string(nc))
for i = 0 to nc - 1
	blob data;
	ulong size
	string ext;
	long pt
	pt = book.getPicture(i,ref data, ref size);
  choose case pt  
		case cxls.PICTURETYPE_PNG  
			ext = "png"; 
		case cxls.PICTURETYPE_JPEG 
			ext = "jpg"; 
		case cxls.PICTURETYPE_GIF  
			ext = "gif"; 
		case cxls.PICTURETYPE_WMF  
			ext = "wmf"; 
		case cxls.PICTURETYPE_EMF  
			ext = "wmf"; 
		case cxls.PICTURETYPE_TIFF 
			ext = "tif";          
	end choose
	integer ifile
	ifile = fileopen(string(i) + "." + ext,streammode!,write!,lockreadwrite!,replace!)
	filewriteex(ifile,data)
	fileclose(ifile)
next
destroy book
end event

type st_1 from statictext within w_examples
integer x = 2455
integer y = 12
integer width = 1294
integer height = 76
integer textsize = -9
integer weight = 700
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 16777215
long backcolor = 134217741
string text = "AutoFilter"
boolean focusrectangle = false
end type

type cb_76 from commandbutton within w_examples
integer x = 2514
integer y = 584
integer width = 997
integer height = 100
integer taborder = 330
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Filter by values"
end type

event clicked;nx_xlbook book
book = create nx_xlbook
nx_xlsheet sheet

sheet = book.addSheet("my");

sheet.writeStr(2, 1, "Country");
sheet.writeStr(2, 2, "Road injures");
sheet.writeStr(2, 3, "Smoking");
sheet.writeStr(2, 4, "Suicide");

sheet.writeStr(3, 1, "USA");     sheet.writeStr(4, 1, "UK"); 
sheet.writeNum(3, 2, 64);         sheet.writeNum(4, 2, 94);
sheet.writeNum(3, 3, 69);         sheet.writeNum(4, 3, 55);
sheet.writeNum(3, 4, 49);         sheet.writeNum(4, 4, 64);

sheet.writeStr(5, 1, "Germany"); sheet.writeStr(6, 1, "Switzerland");
sheet.writeNum(5, 2, 88);         sheet.writeNum(6, 2, 93); 
sheet.writeNum(5, 3, 46);         sheet.writeNum(6, 3, 54);
sheet.writeNum(5, 4, 55);         sheet.writeNum(6, 4, 50);

sheet.writeStr(7, 1, "Spain");   sheet.writeStr(8, 1, "Italy"); 
sheet.writeNum(7, 2, 86);         sheet.writeNum(8, 2, 75); 
sheet.writeNum(7, 3, 47);         sheet.writeNum(8, 3, 52);
sheet.writeNum(7, 4, 69);         sheet.writeNum(8, 4, 71);

sheet.writeStr(9, 1, "Greece");  sheet.writeStr(10, 1, "Japan");
sheet.writeNum(9, 2, 67);         sheet.writeNum(10, 2, 91);
sheet.writeNum(9, 3, 23);         sheet.writeNum(10, 3, 57);
sheet.writeNum(9, 4, 87);         sheet.writeNum(10, 4, 36);

nx_xlautoFilter autoFilter
autoFilter = sheet.autoFilter();
autoFilter.setRef(2, 10, 1, 4);

// filter by countries
nx_xlfiltercolumn column
column = autoFilter.column(0);

column.addFilter("Japan");
column.addFilter("USA");
column.addFilter("Switzerland");

sheet.applyFilter();

book.save("out.xlsx");

destroy book
end event

type cb_75 from commandbutton within w_examples
integer x = 2514
integer y = 464
integer width = 997
integer height = 100
integer taborder = 320
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Number custom filter"
end type

event clicked;nx_xlbook book
book = create nx_xlbook
nx_xlsheet sheet

sheet = book.addSheet("my");

sheet.writeStr(2, 1, "Country");
sheet.writeStr(2, 2, "Road injures");
sheet.writeStr(2, 3, "Smoking");
sheet.writeStr(2, 4, "Suicide");

sheet.writeStr(3, 1, "USA");     sheet.writeStr(4, 1, "UK"); 
sheet.writeNum(3, 2, 64);         sheet.writeNum(4, 2, 94);
sheet.writeNum(3, 3, 69);         sheet.writeNum(4, 3, 55);
sheet.writeNum(3, 4, 49);         sheet.writeNum(4, 4, 64);

sheet.writeStr(5, 1, "Germany"); sheet.writeStr(6, 1, "Switzerland");
sheet.writeNum(5, 2, 88);         sheet.writeNum(6, 2, 93); 
sheet.writeNum(5, 3, 46);         sheet.writeNum(6, 3, 54);
sheet.writeNum(5, 4, 55);         sheet.writeNum(6, 4, 50);

sheet.writeStr(7, 1, "Spain");   sheet.writeStr(8, 1, "Italy"); 
sheet.writeNum(7, 2, 86);         sheet.writeNum(8, 2, 75); 
sheet.writeNum(7, 3, 47);         sheet.writeNum(8, 3, 52);
sheet.writeNum(7, 4, 69);         sheet.writeNum(8, 4, 71);

sheet.writeStr(9, 1, "Greece");  sheet.writeStr(10, 1, "Japan");
sheet.writeNum(9, 2, 67);         sheet.writeNum(10, 2, 91);
sheet.writeNum(9, 3, 23);         sheet.writeNum(10, 3, 57);
sheet.writeNum(9, 4, 87);         sheet.writeNum(10, 4, 36);

nx_xlautoFilter autoFilter
autoFilter = sheet.autoFilter();
autoFilter.setRef(2, 10, 1, 4);

    // shows countries with road injures < 70 or with road injures > 90
nx_xlfiltercolumn column
column = autoFilter.column(1);
column.setCustomFilterex(cxls.OPERATOR_LESS_THAN, "70", cxls.OPERATOR_GREATER_THAN, "90");
autoFilter.setSort(1, true);

sheet.applyFilter();

book.save("out.xlsx");

destroy book
end event

type cb_74 from commandbutton within w_examples
integer x = 2514
integer y = 344
integer width = 997
integer height = 100
integer taborder = 320
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "String custom filter"
end type

event clicked;nx_xlbook book
book = create nx_xlbook
nx_xlsheet sheet

sheet = book.addSheet("my");

sheet.writeStr(2, 1, "Country");
sheet.writeStr(2, 2, "Road injures");
sheet.writeStr(2, 3, "Smoking");
sheet.writeStr(2, 4, "Suicide");

sheet.writeStr(3, 1, "USA");     sheet.writeStr(4, 1, "UK"); 
sheet.writeNum(3, 2, 64);         sheet.writeNum(4, 2, 94);
sheet.writeNum(3, 3, 69);         sheet.writeNum(4, 3, 55);
sheet.writeNum(3, 4, 49);         sheet.writeNum(4, 4, 64);

sheet.writeStr(5, 1, "Germany"); sheet.writeStr(6, 1, "Switzerland");
sheet.writeNum(5, 2, 88);         sheet.writeNum(6, 2, 93); 
sheet.writeNum(5, 3, 46);         sheet.writeNum(6, 3, 54);
sheet.writeNum(5, 4, 55);         sheet.writeNum(6, 4, 50);

sheet.writeStr(7, 1, "Spain");   sheet.writeStr(8, 1, "Italy"); 
sheet.writeNum(7, 2, 86);         sheet.writeNum(8, 2, 75); 
sheet.writeNum(7, 3, 47);         sheet.writeNum(8, 3, 52);
sheet.writeNum(7, 4, 69);         sheet.writeNum(8, 4, 71);

sheet.writeStr(9, 1, "Greece");  sheet.writeStr(10, 1, "Japan");
sheet.writeNum(9, 2, 67);         sheet.writeNum(10, 2, 91);
sheet.writeNum(9, 3, 23);         sheet.writeNum(10, 3, 57);
sheet.writeNum(9, 4, 87);         sheet.writeNum(10, 4, 36);

nx_xlautoFilter autoFilter
autoFilter = sheet.autoFilter();
autoFilter.setRef(2, 10, 1, 4);

// shows all countries with the first letter G
nx_xlfiltercolumn column
column = autoFilter.column(0); 
column.setCustomFilterex(cxls.OPERATOR_EQUAL, "G*");

sheet.applyFilter();

book.save("out.xlsx");

destroy book
end event

type cb_73 from commandbutton within w_examples
integer x = 2514
integer y = 224
integer width = 997
integer height = 100
integer taborder = 310
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Sorting"
end type

event clicked;    nx_xlbook book
	 book = create nx_xlbook
    nx_xlsheet sheet
    sheet = book.addSheet("my");

    sheet.writeStr(2, 1, "Country");
    sheet.writeStr(2, 2, "Road injures");
    sheet.writeStr(2, 3, "Smoking");
    sheet.writeStr(2, 4, "Suicide");

    sheet.writeStr(3, 1, "USA");     sheet.writeStr(4, 1, "UK"); 
    sheet.writeNum(3, 2, 64);         sheet.writeNum(4, 2, 94);
    sheet.writeNum(3, 3, 69);         sheet.writeNum(4, 3, 55);
    sheet.writeNum(3, 4, 49);         sheet.writeNum(4, 4, 64);

    sheet.writeStr(5, 1, "Germany"); sheet.writeStr(6, 1, "Switzerland");
    sheet.writeNum(5, 2, 88);         sheet.writeNum(6, 2, 93); 
    sheet.writeNum(5, 3, 46);         sheet.writeNum(6, 3, 54);
    sheet.writeNum(5, 4, 55);         sheet.writeNum(6, 4, 50);

    sheet.writeStr(7, 1, "Spain");   sheet.writeStr(8, 1, "Italy"); 
    sheet.writeNum(7, 2, 86);         sheet.writeNum(8, 2, 75); 
    sheet.writeNum(7, 3, 47);         sheet.writeNum(8, 3, 52);
    sheet.writeNum(7, 4, 69);         sheet.writeNum(8, 4, 71);

    sheet.writeStr(9, 1, "Greece");  sheet.writeStr(10, 1, "Japan");
    sheet.writeNum(9, 2, 67);         sheet.writeNum(10, 2, 91);
    sheet.writeNum(9, 3, 23);         sheet.writeNum(10, 3, 57);
    sheet.writeNum(9, 4, 87);         sheet.writeNum(10, 4, 36);

	nx_xlautofilter autofilter
    autoFilter = sheet.autoFilter();
    autoFilter.setRef(2, 10, 1, 4);

    // sorting by road injures

    autoFilter.setSort(1, true); 

    sheet.applyFilter();
   
    book.save("out.xlsx");

    destroy book
end event

type cb_72 from commandbutton within w_examples
integer x = 2514
integer y = 104
integer width = 997
integer height = 100
integer taborder = 320
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Top N number of items to filter by"
end type

event clicked;nx_xlbook book
book = create nx_xlbook
nx_xlsheet sheet
sheet = book.addSheet("my")

sheet.writeStr(2, 1, "Country");
sheet.writeStr(2, 2, "Road injures");
sheet.writeStr(2, 3, "Smoking");
sheet.writeStr(2, 4, "Suicide");

sheet.writeStr(3, 1, "USA");     sheet.writeStr(4, 1, "UK"); 
sheet.writeNum(3, 2, 64);         sheet.writeNum(4, 2, 94);
sheet.writeNum(3, 3, 69);         sheet.writeNum(4, 3, 55);
sheet.writeNum(3, 4, 49);         sheet.writeNum(4, 4, 64);

sheet.writeStr(5, 1, "Germany"); sheet.writeStr(6, 1, "Switzerland");
sheet.writeNum(5, 2, 88);         sheet.writeNum(6, 2, 93); 
sheet.writeNum(5, 3, 46);         sheet.writeNum(6, 3, 54);
sheet.writeNum(5, 4, 55);         sheet.writeNum(6, 4, 50);

sheet.writeStr(7, 1, "Spain");   sheet.writeStr(8, 1, "Italy"); 
sheet.writeNum(7, 2, 86);         sheet.writeNum(8, 2, 75); 
sheet.writeNum(7, 3, 47);         sheet.writeNum(8, 3, 52);
sheet.writeNum(7, 4, 69);         sheet.writeNum(8, 4, 71);

sheet.writeStr(9, 1, "Greece");  sheet.writeStr(10, 1, "Japan");
sheet.writeNum(9, 2, 67);         sheet.writeNum(10, 2, 91);
sheet.writeNum(9, 3, 23);         sheet.writeNum(10, 3, 57);
sheet.writeNum(9, 4, 87);         sheet.writeNum(10, 4, 36);

nx_xlautofilter autofilter
autoFilter = sheet.autoFilter();
autoFilter.setRef(2, 10, 1, 4);

// shows 3 countries with the safest roads
nx_xlfiltercolumn column
column = autoFilter.column(1);
column.setTop10(3); 

sheet.applyFilter();

book.save("out.xlsx");

destroy book
end event

type st_6 from statictext within w_examples
integer x = 2752
integer y = 2192
integer width = 1042
integer height = 120
integer textsize = -9
integer weight = 700
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long backcolor = 67108864
string text = "bug report:tbrothe@gmail.com"
alignment alignment = right!
boolean focusrectangle = false
end type

type cb_71 from commandbutton within w_examples
integer x = 59
integer y = 1524
integer width = 974
integer height = 100
integer taborder = 310
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Customizing fonts"
end type

event clicked;nx_xlbook book
book = create nx_xlbook
book.createxls( cxls.type_xlsx)

nx_xlsheet sheet
sheet = book.addSheet("Sheet1")

string fonts[] = {"Aria", "Arial Black", "Comic Sans MS", "Courier New",&
							"Impact", "Times New Roman", "Verdana"}

int i
for i = 1 to upperbound(fonts)
	nx_xlfont font
	font = book.addFont()
	font.setSize(16)
	font.setName(fonts[i])
	nx_xlformat format
	format = book.addFormat()
	format.setFont(font)
	sheet.writeStr(i + 1, 3, fonts[i], format)
next

int fontSize[] = {8, 10, 12, 14, 16, 20, 25}

for i = 1 to upperbound(fontSize) 
	font = book.addFont()
	font.setSize(fontSize[i])        
	format = book.addFormat()
	format.setFont(font)
	sheet.writeStr(i + 1, 7, "Text", format)
next  

font = book.addFont()
font.setSize(16)
format = book.addFormat()        
format.setRotation(255)
format.setFont(font)
sheet.writeStr(2, 9, "Vertica", format)
sheet.setMerge(2, 8, 9, 9)

nx_xlfont boldFont
boldFont = book.addFont()
boldFont.setBold(true)
nx_xlformat boldFormat
boldFormat = book.addFormat()
boldFormat.setFont(boldFont)

nx_xlfont italicFont
italicFont = book.addFont()
italicFont.setItalic(true) 
nx_xlformat italicFormat
italicFormat = book.addFormat()
italicFormat.setFont(italicFont)

nx_xlfont underlineFont
underlineFont = book.addFont()
underlineFont.setUnderline(cxls.UNDERLINE_SINGLE)
nx_xlformat underlineFormat
underlineFormat = book.addFormat()
underlineFormat.setFont(underlineFont)

nx_xlfont strikeoutFont
strikeoutFont = book.addFont()
strikeoutFont.setStrikeOut(true)
nx_xlformat strikeoutFormat
strikeoutFormat = book.addFormat()
strikeoutFormat.setFont(strikeoutFont)

sheet.writeStr(2, 1, "Norma")
sheet.writeStr(3, 1, "Bold", boldFormat)      
sheet.writeStr(4, 1, "Italic", italicFormat)      
sheet.writeStr(5, 1, "Underline", underlineFormat)      
sheet.writeStr(6, 1, "Strikeout", strikeoutFormat)    

if book.save("fonts.xlsx",false) Then
	messagebox('','complete')
end if

destroy book
end event

type cb_70 from commandbutton within w_examples
integer x = 59
integer y = 1404
integer width = 974
integer height = 100
integer taborder = 300
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Aligning, colors and borders"
end type

event clicked;nx_xlbook book
book = create nx_xlbook

book.createxls( cxls.type_xlsx)

nx_xlsheet sheet
sheet = book.addSheet("my")

sheet.setDisplayGridlines(false)

sheet.setCol(1, 1, 30)
sheet.setCol(3, 3, 11.4)
sheet.setCol(4, 4, 2)
sheet.setCol(5, 5, 15)
sheet.setCol(6, 6, 2)
sheet.setCol(7, 7, 15.4)

string nameAlignH[] = {"ALIGNH_LEFT", "ALIGNH_CENTER", "ALIGNH_RIGHT"}
int alignH[] = {cxls.ALIGNH_LEFT, cxls.ALIGNH_CENTER, cxls.ALIGNH_RIGHT}

int i

for i = 1 to upperbound(alignH)
	nx_xlformat format
	format = book.addFormat()
	format.setAlignH(alignH[i])
	format.setBorder(cxls.borderstyle_thin)
	sheet.writeStr(i * 2 + 2, 1, nameAlignH[i], format)
next  

string nameAlignV[] = {"ALIGNV_TOP", "ALIGNV_CENTER", "ALIGNV_BOTTOM"}
long alignV[] = {cxls.ALIGNV_TOP, cxls.ALIGNV_CENTER, cxls.ALIGNV_BOTTOM}

for i = 1 to upperbound(alignV)
	format = book.addFormat()
	format.setAlignV(alignV[i])
	format.setBorder(cxls.borderstyle_thin)
	sheet.writeStr(4, i * 2 + 1, nameAlignV[i], format)
	sheet.setMerge(4, 8, i * 2 + 1, i * 2 + 1)
next

string nameBorderStyle[] = {"BORDERSTYLE_MEDIUM", "BORDERSTYLE_DASHED", &
										 "BORDERSTYLE_DOTTED", "BORDERSTYLE_THICK",& 
										 "BORDERSTYLE_DOUBLE", "BORDERSTYLE_DASHDOT"}
long borderStyle[] = {cxls.BORDERSTYLE_MEDIUM, cxls.BORDERSTYLE_DASHED,cxls.BORDERSTYLE_DOTTED, &
								cxls.BORDERSTYLE_THICK, cxls.BORDERSTYLE_DOUBLE, cxls.BORDERSTYLE_DASHDOT}

for i = 1 to upperbound(nameBorderStyle)       
	format = book.addFormat()
	format.setBorder(borderStyle[i])
	sheet.writeStr(i * 2 + 12, 1, nameBorderStyle[i], format)
next 

string nameColors[] = {"COLOR_RED", "COLOR_BLUE", "COLOR_YELLOW", &
								  "COLOR_PINK", "COLOR_GREEN", "COLOR_GRAY25"}
long colors[] = {cxls.COLOR_RED, cxls.COLOR_BLUE, cxls.COLOR_YELLOW, cxls.COLOR_PINK, cxls.COLOR_GREEN, &
				 cxls.COLOR_GRAY25}
long fillPatterns[] = {cxls.FILLPATTERN_GRAY50, cxls.FILLPATTERN_HORSTRIPE, &
								 cxls.FILLPATTERN_VERSTRIPE, cxls.FILLPATTERN_REVDIAGSTRIPE,&
								 cxls.FILLPATTERN_THINVERSTRIPE, cxls.FILLPATTERN_THINHORCROSSHATCH}


for i = 1 to upperbound(nameColors)
	nx_xlformat format1
	format1 = book.addFormat()
	format1.setFillPattern(cxls.FILLPATTERN_SOLID)
	format1.setPatternForegroundColor(colors[i])
	sheet.writeBlank(i * 2 + 12, 3, format1)
	
	nx_xlformat format2
	format2 = book.addFormat()
	format2.setFillPattern(fillPatterns[i])
	format2.setPatternForegroundColor(colors[i])
	sheet.writeBlank(i * 2 + 12, 5, format2)
	
	nx_xlfont font
	font = book.addFont()
	font.setColor(colors[i])
	
	nx_xlformat format3
	format3 = book.addFormat()
	format3.setBorder(cxls.borderstyle_thin)
	format3.setBorderColor(colors[i])        
	format3.setFont(font)
	sheet.writeStr(i * 2 + 12, 7, nameColors[i], format3)
next


if book.save("acb.xlsx",false) then
	messagebox('','complete')
end if

destroy book
end event

type cb_69 from commandbutton within w_examples
integer x = 59
integer y = 1288
integer width = 974
integer height = 100
integer taborder = 340
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Using number formats"
end type

event clicked;nx_xlbook book
book = create nx_xlbook
book.createxls( cxls.type_xlsx)

nx_xlsheet sheet
sheet = book.addSheet("my")

sheet.setCol(0, 0, 38)
sheet.setCol(1, 1, 10)    


// built-in number formats
nx_xlformat format1,format2,format3,format4,format5,format6,format7,format8,format9,format10,format11,format12

format1 = book.addFormat()
format1.setNumFormat(cxls.NUMFORMAT_NUMBER_D2)

sheet.writeStr(3, 0, "NUMFORMAT_NUMBER_D2")
sheet.writeNum(3, 1, 2.5681, format1)

format2 = book.addFormat()
format2.setNumFormat(cxls.NUMFORMAT_NUMBER_SEP)

sheet.writeStr(4, 0, "NUMFORMAT_NUMBER_SEP")
sheet.writeNum(4, 1, 2500000, format2)

format3 = book.addFormat()
format3.setNumFormat(cxls.NUMFORMAT_CURRENCY_NEGBRA)

sheet.writeStr(5, 0, "NUMFORMAT_CURRENCY_NEGBRA")
sheet.writeNum(5, 1, -500, format3)

format4 = book.addFormat()
format4.setNumFormat(cxls.NUMFORMAT_PERCENT)

sheet.writeStr(6, 0, "NUMFORMAT_PERCENT")
sheet.writeNum(6, 1, -0.25, format4)

format5 = book.addFormat()
format5.setNumFormat(cxls.NUMFORMAT_SCIENTIFIC_D2)

sheet.writeStr(7, 0, "NUMFORMAT_SCIENTIFIC_D2")
sheet.writeNum(7, 1, 890, format5)

format6 = book.addFormat()
format6.setNumFormat(cxls.NUMFORMAT_FRACTION_ONEDIG)

sheet.writeStr(8, 0, "NUMFORMAT_FRACTION_ONEDIG")
sheet.writeNum(8, 1, 0.75, format6)

format7 = book.addFormat()
format7.setNumFormat(cxls.NUMFORMAT_DATE)

sheet.writeStr(9, 0, "NUMFORMAT_DATE")
sheet.writeNum(9, 1, book.datePack(2020, 5, 16,0,0,0,0), format7)

format8 = book.addFormat()
format8.setNumFormat(cxls.NUMFORMAT_CUSTOM_MON_YY)

sheet.writeStr(10, 0, "NUMFORMAT_CUSTOM_MON_YY")
sheet.writeNum(10, 1, book.datePack(2020, 5, 16,0,0,0,0), format8)

// custom number formats

format9 = book.addFormat()
format9.setNumFormat(book.addCustomNumFormat("#.###"))    

sheet.writeStr(12, 0, "#.###")
sheet.writeNum(12, 1, 20.5627, format9)

format10 = book.addFormat()
format10.setNumFormat(book.addCustomNumFormat("#.00"))    

sheet.writeStr(13, 0, "#.00")
sheet.writeNum(13, 1, 4.8, format10)

format11 = book.addFormat()
format11.setNumFormat(book.addCustomNumFormat('0.00 "dollars"'))    

sheet.writeStr(14, 0, '0.00 "dollars"')
sheet.writeNum(14, 1, 1.23, format11)

format12 = book.addFormat()
format12.setNumFormat(book.addCustomNumFormat("[Red][<=100];[Green][>100]"))    

sheet.writeStr(15, 0, "[Red][<=100];[Green][>100]")
sheet.writeNum(15, 1, 60, format12)



if book.save( "numformats.xlsx"/*string filename*/,false /*boolean usetempfile */) then
	messagebox('','complete')
end if
destroy book
end event

type cb_68 from commandbutton within w_examples
integer x = 59
integer y = 1172
integer width = 974
integer height = 100
integer taborder = 330
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Inserting rows and columns"
end type

event clicked;nx_xlbook book
book = create nx_xlbook
book.createxls( cxls.type_xlsx)


nx_xlsheet sheet
sheet = book.addSheet("Sheet1")

nx_xlformat format
format = book.addformat()

long row,col
for row = 1 to 30
  for col = 0 to 10
		sheet.writeNum(row, col,rand(10),format)
	next
next

sheet.insertRow(5, 10,false)
sheet.insertRow(20, 22,false)

sheet.insertCol(4, 5,false)
sheet.insertCol(8, 8,false)


if book.save( "insert.xlsx"/*string filename*/,false /*boolean usetempfile */) then
	messagebox('','complete')
end if
destroy book
end event

type cb_67 from commandbutton within w_examples
integer x = 59
integer y = 940
integer width = 974
integer height = 100
integer taborder = 320
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Merging cells"
end type

event clicked;nx_xlbook book
book = create nx_xlbook
book.createxls( cxls.type_xlsx)

nx_xlformat format
format = book.addFormat();
format.setAlignH(cxls.ALIGNH_CENTER);
format.setAlignV(cxls.ALIGNV_CENTER);

nx_xlsheet sheet
sheet = book.addSheet("Sheet1")

sheet.writeStr(3, 1, "Hello World !", format)

sheet.setMerge(3, 5, 1, 5)

sheet.setMerge(7, 20, 1, 2)
sheet.setMerge(7, 20, 4, 5)

sheet.writeNum(7, 1, 1, format)
sheet.writeNum(7, 4, 2, format)


if book.save( "merge.xlsx"/*string filename*/,false /*boolean usetempfile */) then
	messagebox('','complete')
end if
destroy book
end event

type cb_66 from commandbutton within w_examples
integer x = 59
integer y = 824
integer width = 974
integer height = 100
integer taborder = 310
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Access to sheet by name"
end type

event clicked;nx_xlbook book
book = Create nx_xlbook

long sheetcount,sheetindex
string sheetname
if book.load( "data.xlsx") Then
	sheetcount = book.getsheetcount( )
	for sheetindex = 0 to sheetcount  - 1	
		sheetname = book.getsheetname( sheetindex)
		messagebox(string(sheetindex),sheetname)
		if sheetname = "mysheetname" Then
			/*get your sheetbyname*/
			/*
				sheet  = book.getsheet( sheetindex)
			*/
		end if
	next
end if
end event

type cb_65 from commandbutton within w_examples
integer x = 59
integer y = 700
integer width = 1029
integer height = 100
integer taborder = 300
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Reading and writing date/time values"
end type

event clicked;nx_xlbook book
book = create nx_xlbook
book.createxls( cxls.type_xlsx)

nx_xlformat format1,format2,format3,format4
format1 = book.addFormat()
format1.setNumFormat(cxls.NUMFORMAT_DATE)

format2 = book.addFormat()
format2.setNumFormat(cxls.NUMFORMAT_CUSTOM_MDYYYY_HMM)

format3 = book.addFormat()
format3.setNumFormat(book.addCustomNumFormat("d mmmm yyyy"))

format4 = book.addFormat()
format4.setNumFormat(cxls.NUMFORMAT_CUSTOM_HMM_AM)

nx_xlsheet sheet
sheet = book.addSheet("Sheet1")


sheet.setCol(1, 1, 15)

// writing

sheet.writeNum(2, 1, book.datePack(2010, 3, 11,0,0,0,0), format1)       
sheet.writeNum(3, 1, book.datePack(2010, 3, 11, 10, 25, 55,0), format2)
sheet.writeNum(4, 1, book.datePack(2010, 3, 11,0,0,0,0), format3)       
sheet.writeNum(5, 1, book.datePack(2010, 3, 11, 10, 25, 55,0), format4)

// reading

long year, month, day , hour, min, sec,msec

book.dateUnpack(sheet.readNum(3, 1), ref year, ref month, ref day, ref hour, ref min, ref sec , ref msec)

messagebox('datetime',			"year:" 		+ string(year) + "~n" + &
										"month:" 	+ string(month) +  "~n" + &
										"day:" 		+ string(day) +  "~n" + &
										"hour:"		+ string(hour) +  "~n" + &
										"min:"	 	+ string(min) +  "~n" + &
										"second:" 	+ string(sec) +  "~n" + &
										"msec:"	 	+ string(msec) )	


if book.save("datetime.xlsx",false) then
	Messagebox('','complete')
end if
destroy book
end event

type cb_64 from commandbutton within w_examples
integer x = 59
integer y = 584
integer width = 974
integer height = 100
integer taborder = 290
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Writing formulas"
end type

event clicked;nx_xlbook book
book = create nx_xlbook
book.createxls( cxls.type_xlsx)

nx_xlformat alFormat
alFormat = book.addFormat()
alFormat.setAlignH(cxls.ALIGNH_LEFT)

nx_xlformat arformat
arFormat = book.addFormat()
arFormat.setAlignH(cxls.ALIGNH_RIGHT)

nx_xlformat alignDateFormat
alignDateFormat = book.addFormat(alFormat)
alignDateFormat.setNumFormat(cxls.NUMFORMAT_DATE)

nx_xlfont linkFont
linkFont = book.addFont()
linkFont.setColor(cxls.COLOR_BLUE)
linkFont.setUnderline(cxls.UNDERLINE_SINGLE)

nx_xlformat linkFormat
linkFormat = book.addFormat(alFormat)
linkFormat.setFont(linkFont)

nx_xlsheet sheet
sheet = book.addSheet("Sheet1")


sheet.setCol(0, 0, 27)
sheet.setCol(1, 1, 10)

sheet.writeNum(2, 1, 40, alFormat)
sheet.writeNum(3, 1, 30, alFormat)
sheet.writeNum(4, 1, 50, alFormat)

sheet.writeStr(6, 0, "SUM(B3:B5) = ", arFormat)        
sheet.writeFormula(6, 1, "SUM(B3:B5)", alFormat)        
sheet.writeStr(7, 0, "AVERAGE(B3:B5) = ", arFormat)        
sheet.writeFormula(7, 1, "AVERAGE(B3:B5)", alFormat)        
sheet.writeStr(8, 0, "MAX(B3:B5) = ", arFormat)        
sheet.writeFormula(8, 1, "MAX(B3:B5)", alFormat)        
sheet.writeStr(9, 0, "MIX(B3:B5) = ", arFormat)        
sheet.writeFormula(9, 1, "MIN(B3:B5)", alFormat)
sheet.writeStr(10, 0, "COUNT(B3:B5) = ", arFormat)      
sheet.writeFormula(10, 1, "COUNT(B3:B5)", alFormat)

sheet.writeStr(12, 0, 'IF(B7 > 100;"large";"small") = ', arFormat)      
sheet.writeFormula(12, 1, 'IF(B7 > 100;"large";"small")', alFormat)

sheet.writeStr(14, 0, "SQRT(25) = ", arFormat)      
sheet.writeFormula(14, 1, "SQRT(25)", alFormat)
sheet.writeStr(15, 0, "RAND() = ", arFormat)      
sheet.writeFormula(15, 1, "RAND()", alFormat)
sheet.writeStr(16, 0, "2*PI() = ", arFormat)      
sheet.writeFormula(16, 1, "2*PI()", alFormat)

sheet.writeStr(18, 0, 'UPPER("libxl") = ', arFormat)      
sheet.writeFormula(18, 1, 'UPPER("libxl")', alFormat)
sheet.writeStr(19, 0, 'LEFT("window";3) = ', arFormat)      
sheet.writeFormula(19, 1, 'LEFT("window";3)', alFormat)
sheet.writeStr(20, 0, 'LEN("string") = ', arFormat)      
sheet.writeFormula(20, 1, 'LEN("string")', alFormat)

sheet.writeStr(22, 0, "DATE(2010;3;11) = ", arFormat)      
sheet.writeFormula(22, 1, "DATE(2010;3;11)", alignDateFormat)
sheet.writeStr(23, 0, "DAY(B23) = ", arFormat)      
sheet.writeFormula(23, 1, "DAY(B23)", alFormat)
sheet.writeStr(24, 0, "MONTH(B23) = ", arFormat)      
sheet.writeFormula(24, 1, "MONTH(B23)", alFormat)
sheet.writeStr(25, 0, "YEAR(B23) = ", arFormat)      
sheet.writeFormula(25, 1, "YEAR(B23)", alFormat)
sheet.writeStr(26, 0, "DAYS360(B23;TODAY()) = ", arFormat)      
sheet.writeFormula(26, 1, "DAYS360(B23;TODAY())", alFormat)

sheet.writeStr(28, 0, "B3+100*(2-COS(0)) = ", arFormat)      
sheet.writeFormula(28, 1, "B3+100*(2-COS(0))", alFormat)
sheet.writeStr(29, 0, "ISNUMBER(B29) = ", arFormat)      
sheet.writeFormula(29, 1, "ISNUMBER(B29)", alFormat)
sheet.writeStr(30, 0, "AND(1;0) = ", arFormat)      
sheet.writeFormula(30, 1, "AND(1;0)", alFormat)

sheet.writeStr(32, 0, "HYPERLINK() = ", arFormat)
sheet.writeFormula(32, 1, 'HYPERLINK("http://www.libxl.com")', linkFormat)

if book.save("formula.xlsx",false) then
		messagebox('','complete')
end if	
destroy book
end event

type cb_63 from commandbutton within w_examples
integer x = 59
integer y = 356
integer width = 974
integer height = 100
integer taborder = 280
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Placing pictures"
end type

event clicked;boolean lb
nx_xlbook book
book = Create nx_xlbook
lb = book.createxls( cxls.TYPE_XLSX)
messagebox('',string(lb))
long id
if lb Then
	id = book.addpicture( "1.jpg")
	messagebox("picid",String(id))
	if id = -1 Then
		messagebox('','picture not found')
		return
	end if
	nx_xlsheet sheet
	sheet = book.addsheet( "sheet1")
	sheet.setpicture( 10/*long row*/,1 /*long col*/, id/*long pictureid*/,1 /*double scale*/,0 /*long offset_x*/,0 /*long offset_y */)

	if book.save( "Placing pictures.xlsx",false /*boolean usetempfile */) Then
		messagebox('','complete')
	end if	
end if
destroy book
end event

type cb_56 from commandbutton within w_examples
integer x = 59
integer y = 228
integer width = 974
integer height = 100
integer taborder = 270
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "reading data"
end type

event clicked;nx_xlbook book
book = Create nx_xlbook
string sout
sout = ''
long row,col,rowlast,collast
long celltype
if book.load( "data.xlsx") Then 
	nx_xlsheet sheet
	sheet = book.getsheet( 2)
	rowlast = sheet.lastrow( )
	collast = sheet.lastcol( )
	for row = 0 to rowlast - 1
		for col = 0 to collast - 1
			celltype = sheet.celltype(row,col)
			sout += "(" + string(row) + "," + string(col) + ") = " 
			if sheet.isformula( row,col /*long col */) Then
				sout += sheet.readformula( row,col /*long col */) + "[formula]~n"
			else
				choose case celltype
					case cxls.celltype_empty
						sout += "[empty]~n"
					case cxls.celltype_number
						sout += string(sheet.readnum( row,col /*long col */)) + "[number]~n"
					case cxls.celltype_datetime	/*use dateunpack*/
						sout += string(sheet.readnum( row,col /*long col */)) + "[date]~n"
					case cxls.celltype_string
						sout += sheet.readstr( row,col /*long col */) + "[string]~n"
					case cxls.celltype_boolean
						sout += string(sheet.readbool( row,col /*long col */)) + "[boolean]~n"	
					case cxls.celltype_blank
						sout += "[blank]~n"
					case cxls.celltype_error
						sout += "[error]~n"
				end choose
			end if
		next
	next	
end if

messagebox('reading data',sout)
destroy book
end event

type cb_55 from commandbutton within w_examples
integer x = 59
integer y = 108
integer width = 974
integer height = 100
integer taborder = 260
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "writing data"
end type

event clicked;nx_xlbook book
book = Create nx_xlbook

book.createxls( cxls.TYPE_XLS)
//
int logoid
logoid = book.addpicture( "logo.png")
//
////fonts
nx_xlfont textfont
textfont = book.addfont( )
textfont.setsize( 8)
textfont.setname( "Arial")

nx_xlfont titlefont
titlefont = book.addfont( textfont )
titlefont.setsize( 38)
titlefont.setcolor( cxls.COLOR_GRAY25)

nx_xlfont font12,font10
font12 = book.addfont( textfont )
font10 = book.addfont( textfont)
font12.setsize( 12)
font10.setsize( 10)

////format
nx_xlformat textformat
textformat = book.addformat( )
textformat.setfont( textfont)
textformat.setalignh( cxls.ALIGNH_LEFT)

nx_xlformat titleformat
titleformat = book.addformat( )
titleformat.setfont( titlefont)
titleformat.setalignh( cxls.ALIGNH_RIGHT)

nx_xlformat companyformat
companyformat = book.addformat( )
companyformat.setfont( font12)

nx_xlformat dateformat
dateformat = book.addformat(textformat)
dateformat.setnumformat( book.addcustomnumformat( "[$-409]mmmm\ d\,\ yyyy;@"))

nx_xlformat phoneformat
phoneformat = book.addformat(textformat)
phoneformat.setnumformat( book.addcustomnumformat( "[<=9999999]###\-####;\(###\)\ ###\-####"))

nx_xlformat borderformat
borderformat = book.addformat(textformat )
borderformat.setborder( cxls.BORDERSTYLE_THIN)
borderformat.setbordercolor( cxls.COLOR_GRAY25)
borderformat.setalignv( cxls.ALIGNV_CENTER)

nx_xlformat percentformat
percentformat = book.addformat(borderformat )
percentformat.setnumformat( book.addcustomnumformat( "#%_)"))
percentformat.setalignh( cxls.ALIGNH_RIGHT)

nx_xlformat textrightformat
textrightformat = book.addformat( textformat )
textRightFormat.setAlignH(cxls.ALIGNH_RIGHT);
textRightFormat.setAlignV(cxls.ALIGNV_CENTER);

nx_xlformat thankformat
thankformat = book.addformat( )
thankFormat.setFont(font10);
thankFormat.setAlignH(cxls.ALIGNH_CENTER);

nx_xlformat dollarformat
dollarformat = book.addformat( borderformat )
dollarformat.setnumformat( book.addcustomnumformat( "_($* # ##0.00_);_($* (# ##0.00);_($* -??_);_(@_)"))
//actions
nx_xlsheet sheet
sheet = book.addsheet( "Sales Receipt±í")

//TODO
sheet.setdisplaygridlines(false)

sheet.setCol(1, 1, 36)
sheet.setCol(0, 0, 10)
sheet.setCol(2, 4, 11)

sheet.setRow(2, 47.25)
sheet.writeStr(2, 1, "Sales Receipt", titleFormat)
sheet.setMerge(2, 2, 1, 4)
sheet.setPicture(2, 1, logoId,1.0,0,0)


sheet.writeStr(4, 0, "Apricot Ltd.", companyFormat)
sheet.writeStr(4, 3, "Date:", textFormat)

//TODO
sheet.writeFormula(4, 4, "TODAY()", dateFormat)

sheet.writeStr(5, 3, "Receipt #:", textFormat)

//TODO
sheet.writeNum(5, 4, 652, textFormat)

sheet.writeStr(8, 0, "Sold to:", textFormat)
sheet.writeStr(8, 1, "John Smith", textFormat)
sheet.writeStr(9, 1, "Pineapple Ltd.", textFormat)
sheet.writeStr(10, 1, "123 Dreamland Street", textFormat)
sheet.writeStr(11, 1, "Moema, 52674", textFormat)

//TODO
sheet.writeNum(12, 1, 2659872055, phoneFormat)

sheet.writeStr(14, 0, "Item #", textFormat)
sheet.writeStr(14, 1, "Description", textFormat)
sheet.writeStr(14, 2, "Qty", textFormat)
sheet.writeStr(14, 3, "Unit Price", textFormat)
sheet.writeStr(14, 4, "Line Total", textFormat)

int row,col
string s
for row = 15 to 37
	sheet.setRow(row, 15)
	for col = 0 to 2
		//TODO
		sheet.writeBlank(row, col, borderFormat)
	next
	//TODO
	sheet.writeBlank(row, 3, dollarFormat)

	//TODO
	//s = sprintf('IF(C{1}>0;ABS(C{2}*D{3});"")',row + 1,row + 1,row + 1)
	s = 'IF(C' +string(row + 1)+ '>0;ABS(C' +string(row + 1)+ '*D' +string(row + 1)+ ');"")'
	sheet.writeFormula(row, 4, s, dollarFormat)
next 



sheet.writeStr(38, 3, "Subtotal ", textRightFormat)
sheet.writeStr(39, 3, "Sales Tax ", textRightFormat)
sheet.writeStr(40, 3, "Total ", textRightFormat)
sheet.writeFormula(38, 4, "SUM(E16:E38)", dollarFormat)
sheet.writeNum(39, 4, 0.2, percentFormat)
sheet.writeFormula(40, 4, "E39+E39*E40", dollarFormat)
sheet.setRow(38, 15)
sheet.setRow(39, 15)
sheet.setRow(40, 15)

sheet.writeStr(42, 0, "Thank you for your business!", thankFormat)
sheet.setMerge(42, 42, 0, 4)

// items

sheet.writeNum(15, 0, 45, borderFormat)
sheet.writeStr(15, 1, "Grapes", borderFormat)
sheet.writeNum(15, 2, 250, borderFormat)
sheet.writeNum(15, 3, 4.5, dollarFormat)

sheet.writeNum(16, 0, 12, borderFormat)
sheet.writeStr(16, 1, "Bananas", borderFormat)
sheet.writeNum(16, 2, 480, borderFormat)
sheet.writeNum(16, 3, 1.4, dollarFormat)

sheet.writeNum(17, 0, 19, borderFormat)
sheet.writeStr(17, 1, "Apples", borderFormat)
sheet.writeNum(17, 2, 180, borderFormat)
sheet.writeNum(17, 3, 2.8, dollarFormat)

book.save("receipt.xls",false)
//book->release();
book.close( )
destroy book

Messagebox('','complete!')
end event

type st_5 from statictext within w_examples
integer x = 9
integer y = 12
integer width = 1509
integer height = 76
integer textsize = -9
integer weight = 700
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 16777215
long backcolor = 134217741
string text = "example"
boolean focusrectangle = false
end type


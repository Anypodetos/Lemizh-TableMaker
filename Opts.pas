unit Opts;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, StdCtrls, Buttons, Menus, IniFiles, MyUtils, Math, ComCtrls,
  FileCtrl, ExtCtrls, Spin, Gauges;

type
  TOptions = class(TForm)
    OKBtn: TBitBtn;
    CancelBtn: TBitBtn;
    PageControl: TPageControl;
    CaseSheet: TTabSheet;
    FileSheet: TTabSheet;
    CaseBox: TGroupBox;
    CaseGrid: TStringGrid;
    SymbolBox: TGroupBox;
    SymbGrid: TStringGrid;
    AddBtn: TBitBtn;
    DelBtn: TBitBtn;
    UpBtn: TBitBtn;
    DownBtn: TBitBtn;
    GroupBox1: TGroupBox;
    DirLabel1: TLabel;
    BackupFilesBox: TCheckBox;
    GroupBox2: TGroupBox;
    DirLabel2: TLabel;
    Label1: TLabel;
    ChangesPanel: TPanel;
    Label3: TLabel;
    ChangesBox: TListBox;
    DuplicatesPanel: TPanel;
    Label4: TLabel;
    DuplicatesBox: TListBox;
    MenuBreakBox: TCheckBox;
    DictSheet: TTabSheet;
    GroupBox3: TGroupBox;
    DirLabel3: TLabel;
    SubstStemsBox: TCheckBox;
    GroupBox4: TGroupBox;
    Dictionary: TListBox;
    UnconfirmedBox: TCheckBox;
    EntriesLabel: TLabel;
    DirectoryListBox1: TDirectoryListBox;
    DirectoryListBox2: TDirectoryListBox;
    DirectoryListBox3: TDirectoryListBox;
    DriveComboBox1: TDriveComboBox;
    DriveComboBox2: TDriveComboBox;
    DriveComboBox3: TDriveComboBox;
    LemBox: TCheckBox;
    GroupBox5: TGroupBox;
    GlossingBox: TCheckBox;
    Bevel2: TBevel;
    UpdateHTMLsBtn: TBitBtn;
    DirPopup: TPopupMenu;
    Newfolder1: TMenuItem;
    NeogrammBtn: TBitBtn;
    NewFolderBtn1: TBitBtn;
    NewFolderBtn2: TBitBtn;
    NewFolderBtn3: TBitBtn;
    FileListBox3: TFileListBox;
    GlossedSheet: TTabSheet;
    GroupBox6: TGroupBox;
    UnglossedLabel: TLabel;
    UnglossedMemo: TMemo;
    GlossedMemo: TMemo;
    UnglossedInclText: TCheckBox;
    PrefsSheet: TTabSheet;
    GroupBox7: TGroupBox;
    Label8: TLabel;
    HintHideEdit: TSpinEdit;
    Label9: TLabel;
    LemFontCorrEdit: TSpinEdit;
    Label11: TLabel;
    Label10: TLabel;
    Label12: TLabel;
    SolvSizeEdit: TSpinEdit;
    Label14: TLabel;
    MemoSizeEdit: TSpinEdit;
    Label13: TLabel;
    Label15: TLabel;
    NgrVerLabel: TLabel;
    Label16: TLabel;
    Appendices: TMemo;
    Bevel3: TBevel;
    Bevel4: TBevel;
    Label17: TLabel;
    TextTags: TMemo;
    Label19: TLabel;
    Label20: TLabel;
    FindPanel: TPanel;
    Label21: TLabel;
    FindDown: TSpeedButton;
    FindClear: TSpeedButton;
    FindEdit: TEdit;
    AppendicesUpTo: TLabel;
    UnglossedPanel: TPanel;
    UnglossedGauge: TGauge;
    UnglossedTimer: TTimer;
    StatSheet: TTabSheet;
    GroupBox8: TGroupBox;
    StatListBox: TListBox;
    StatInclText: TCheckBox;
    GroupBox9: TGroupBox;
    UploadCheckBox: TCheckBox;
    Label32: TLabel;
    HostEdit: TEdit;
    Label33: TLabel;
    UserEdit: TEdit;
    Label34: TLabel;
    PasswordEdit: TEdit;
    Label35: TLabel;
    UploadPathEdit: TEdit;
    ShowUploadCheckBox: TCheckBox;
    UploadSolvCheckBox: TCheckBox;
    Label2: TLabel;
    UploadSolvPathEdit: TEdit;
    Label5: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure AddBtnClick(Sender: TObject);
    procedure UpDownBtnClick(Sender: TObject);
    procedure DelBtnClick(Sender: TObject);
    procedure SymbGridClick(Sender: TObject);
    procedure CaseGridKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure GridSetEditText(Sender: TObject; ACol, ARow: Integer; const Value: string);
    procedure SymbGridKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure SymbGridKeyPress(Sender: TObject; var Key: Char);
    procedure UpdateDict(Sender: TObject);
    procedure DictionaryDrawItem(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
    procedure NeogrammBtnClick(Sender: TObject);
    procedure LemBoxClick(Sender: TObject);
    procedure UpdateHTMLsBtnClick(Sender: TObject);
    procedure DirectoryListBox3Change(Sender: TObject);
    procedure Newfolder1Click(Sender: TObject);
    procedure PageControlChange(Sender: TObject);
    procedure FindClearClick(Sender: TObject);
    procedure FindEditChange(Sender: TObject);
    procedure FindDownClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UnglossedLabelMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure UnglossedTimerTimer(Sender: TObject);
    procedure StatListBoxDrawItem(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
    procedure UploadCheckBoxClick(Sender: TObject);
  private
    DictWriteTime: TDateTime;
  public
    procedure UpdateButtons;
  end;

var
  Options: TOptions;

function SubstituteStems(S: string; Lemizh: Boolean; Dict: TStrings): string;

implementation

uses LTableMaker;

{$R *.DFM}

procedure TOptions.FormCreate(Sender: TObject);
var i: Integer;
begin
  Constraints.MinHeight:=Height div 2;
  Constraints.MinWidth:=Width;
  CaseGrid.ColWidths[2]:=CaseGrid.Width-2*CaseGrid.DefaultColWidth-30;
  CaseGrid.Cells[0,0]:='Ending';
  CaseGrid.Cells[1,0]:='Gloss';
  CaseGrid.Cells[2,0]:='Name';
  for i:=0 to High(TableMaker.CaseAbbrs) do CaseGrid.Cells[0,i+1]:=CaseEnding(i);
  SymbGrid.ColWidths[2]:=SymbGrid.Width-2*SymbGrid.DefaultColWidth-30;
  SymbGrid.Cells[0,0]:='Code';
  SymbGrid.Cells[1,0]:='Output';
  SymbGrid.Cells[2,0]:='Name';
  AppendicesUpTo.Caption:='(up to '+IntToStr(MaxAppendices)+'):';
end; {FormCreate}

{------------------------------------------------ Cases / Symbols -----------------------------------------------------------}

procedure TOptions.CaseGridKeyPress(Sender: TObject; var Key: Char);
begin
  if CaseGrid.Col=1 then
    if Key in ['A'..'Z'] then Key:=LowerCase(Key)[1] else if not (Key in [#0..#31, 'a'..'z']) then begin
      Key:=#0;
      MessageBeep(MB_IconError);
    end;
end; {CaseGridKeyPress}

procedure TOptions.SymbGridClick(Sender: TObject);
begin
  UpdateButtons;
end; {SymbGridClick}

procedure TOptions.AddBtnClick(Sender: TObject);
begin
  if SymbGrid.RowCount>=250 then TangoMessageBox('The maximum number of symbols is 250.', mtError, [mbOK], '') else with SymbGrid do begin
    RowCount:=RowCount+1;
    Col:=0;
    Row:=SymbGrid.RowCount-1;
    EditorMode:=True;
    SetFocus;
    UpdateButtons;
  end;
end; {AddBtnClick}

procedure TOptions.DelBtnClick(Sender: TObject);
var i, j: Integer;
begin
  with SymbGrid do for i:=0 to 2 do for j:=Row to RowCount-2 do Cells[i,j]:=Cells[i,j+1];
  for i:=0 to 2 do SymbGrid.Cells[i, SymbGrid.RowCount-1]:='';
  SymbGrid.RowCount:=Max(SymbGrid.RowCount-1, 2);
  GridSetEditText(nil, SymbGrid.Col, SymbGrid.Row, '');
  UpdateButtons;
end; {DelBtnClick}

procedure TOptions.UpDownBtnClick(Sender: TObject);
var i: Integer;
begin
  with SymbGrid do for i:=0 to 2 do Cols[i].Exchange(Row, Row+TBitBtn(Sender).Tag);
  SymbGrid.Row:=SymbGrid.Row+TBitBtn(Sender).Tag;
  GridSetEditText(nil, SymbGrid.Col, SymbGrid.Row, '');
end; {UpDownBtnClick}

procedure TOptions.UpdateButtons;
begin
  UpBtn.Enabled:=SymbGrid.Row>1;
  DownBtn.Enabled:=SymbGrid.Row<SymbGrid.RowCount-1;
end; {UpdateButtons}

procedure TOptions.FormClose(Sender: TObject; var Action: TCloseAction);
var i, j: Integer;
    stc, sts: string;
    bs, bh: Boolean;
begin
  if ModalResult=mrOK then begin
    stc:='';
    with CaseGrid do for i:=1 to RowCount-1 do for j:=1 to Length(Cells[1,i]) do if not (Cells[1,i][j] in ['a'..'z']) then begin
      stc:=stc+Cells[1,i][j];
      Row:=i;
    end;
    with SymbGrid do for i:=1 to RowCount-1 do for j:=1 to Length(Cells[0,i]) do if Cells[0,i][j] in ['(', ')', '[', ']', '{', '}'] then begin
      sts:=sts+Cells[0,i][j];
      Row:=i;
    end;
    bs:=False;  bh:=False;
    with SymbGrid do for i:=1 to RowCount-1 do if (Cells[0,i]<>'') and (MultPos(['!'..',', '.'..'@' , '['..'`', '{'..'~', 'Ä'..'ø', '◊', '˜'], Cells[0,i])=NoPos) then begin
      bs:=True;
      if Pos('-', Cells[0,i])>0 then bh:=True;
      Row:=i;
    end;
    if (stc<>'') or (sts<>'') or bs then begin
      TangoMessageBox(IfThen(stc<>'', 'Invalid character'+Copy('s', 1, Ord(Length(stc)>1))+' in case glosses: '+stc+#13#10)
                     +IfThen(sts<>'', 'Brackets not allowed in symbols: '+sts+#13#10)
                     +IfThen(bs, 'Symbols must contain at least one non-alphabetic character'+IfThen(bh, ' (excluding the hyphen)')), mtError, [mbOK], '');
      ModalResult:=mrNone;
    end else
      with SymbGrid do for i:=1 to RowCount-1 do if Cells[1,i]='' then Cells[1,i]:=Cells[0,i];
  end;
end; {FormClose}

procedure TOptions.SymbGridKeyPress(Sender: TObject; var Key: Char);
begin
  if CaseGrid.Col=1 then
    if Key in ['(', ')', '[', ']', '{', '}'] then begin
      Key:=#0;
      MessageBeep(MB_IconError);
    end;
end; {SymbGridKeyPress}

procedure TOptions.SymbGridKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key=VK_Down) and (Shift=[]) and (SymbGrid.Row=SymbGrid.RowCount-1) then AddBtnClick(nil);
end; {SymbGridKeyDown}

procedure TOptions.GridSetEditText(Sender: TObject; ACol, ARow: Integer; const Value: string);
  procedure AddIfThere(st: string);
  var t, l, r, c, n: Integer;
  begin
    n:=0;
    if (st<>'') and (st<>'-') then with TableMaker do begin
      for t:=0 to TabControl.Tabs.Count-2 do for l:=0 to Length(Tables[t])-1 do for r:=0 to Length(Tables[t,l].HTMLs)-1 do
        for c:=0 to ColCounts[t]-1+Ord(t=SolvTab) do if Pos(st, Tables[t,l].HTMLs[r,c])>0 then Inc(n);
      for t:=0 to Length(TextPages)-1 do if Pos(st, TextPages[t].Text)>0 then Inc(n);
    end;
    if n>0 then ChangesBox.Items.Add(st+' ('+IntToStr(n)+')');
  end;
var i, j: Integer;
    b: Boolean;
    old: set of Byte;
begin
  ChangesBox.Items.Clear;
  DuplicatesBox.Items.Clear;
  with CaseGrid do for i:=0 to High(TableMaker.CaseAbbrs) do begin
    if Cells[1,i+1]<>TableMaker.CaseAbbrs[i] then begin
      AddIfThere('-'+Cells[1,i+1]);
      AddIfThere('-'+TableMaker.CaseAbbrs[i]);
    end;
    for j:=0 to High(TableMaker.CaseAbbrs) do if (i<>j) and ((Pos(Cells[1,i+1], Cells[1,j+1])=1) and not ((j<i) and (Cells[1,i+1]=Cells[1,j+1]))) then
      DuplicatesBox.Items.Add('-'+Cells[1,i+1]+' < -'+Cells[1,j+1]);
  end;
  old:=[];
  with SymbGrid do for i:=1 to RowCount-1 do if Cells[0,i]<>'' then begin
    b:=True;
    for j:=0 to i-1 do if Pos(Cells[0,i], Cells[0,j])>0 then b:=False;
    for j:=0 to Length(TableMaker.Symbols[0])-1 do if Cells[0,i]=TableMaker.Symbols[0,j] then begin
      b:=False;
      old:=old+[j];
    end;
    if b then AddIfThere(Cells[0,i]);
    for j:=i+1 to RowCount-1 do if Pos(Cells[0,i], Cells[0,j])>0 then
      DuplicatesBox.Items.Add(Cells[0,i]+' vs. '+Cells[0,j]);
  end;
  for i:=0 to Length(TableMaker.Symbols[0])-1 do if (TableMaker.Symbols[0,i]<>'') and not (i in old) then AddIfThere(TableMaker.Symbols[0,i]);
  ChangesPanel.Visible:=ChangesBox.Items.Count>0;
  DuplicatesPanel.Visible:=DuplicatesBox.Items.Count>0;
end; {GridSetEditText}

{------------------------------------------------ Files -----------------------------------------------------------}

procedure TOptions.UpdateHTMLsBtnClick(Sender: TObject);
var rec: TSearchRec;
begin
  if TangoMessageBox('Delete and replace all files in '+DirectoryListBox2.Directory+'?', mtWarning, [mbYes, mbNo], '')=idYes then begin
    UpdateHTMLsBtn.Enabled:=False;
    if FindFirst(DirectoryListBox2.Directory+'\*.*', 0, rec)=0 then begin
      DeleteFile(rec.Name);
      while FindNext(rec)=0 do DeleteFile(rec.Name);
    end;
    FindClose(rec);
    TableMaker.SaveSolvHTMLs(True);
    UpdateHTMLsBtn.Enabled:=True;
  end;
end; {UpdateHTMLsBtnClick}

procedure TOptions.Newfolder1Click(Sender: TObject);
var dlb: TComponent;
    st: string;
begin
  if Sender is TPopupMenu then dlb:=TDirectoryListBox(DirPopup.PopupComponent) else dlb:=FindComponent('DirectoryListBox'+IntToStr(TComponent(Sender).Tag));
  with TDirectoryListBox(dlb) do if TangoInputQuery('New subfolder', 'of '+Directory+':', st) then try
    MkDir(Directory+'\'+st);
    Update;
  except TangoMessageBox('Folder has an invalid name or already exists.', mtError, [mbOK], '') end;
end; {Newfolder1Click}

procedure TOptions.UploadCheckBoxClick(Sender: TObject);
begin
  with TEdit(FindComponent('Upload'+IfThen(TCheckBox(Sender).Tag=1, 'Solv')+'PathEdit')) do if TCheckBox(Sender).Checked then Color:=clWindow else Color:=clBtnFace;
end;

{------------------------------------------------ Dictionary -----------------------------------------------------------}

procedure TOptions.DirectoryListBox3Change(Sender: TObject);
begin
  SubstStemsBox.Checked:=True;
  UpdateDict(nil);
end; {DirectoryListBox3Change}

procedure TOptions.UpdateDict(Sender: TObject);
const lModLem = 6;
      letters = 'aeyioOuUjczvwxhsqfgdbktpnmrRl';
      openbrackets = '([{';  closebrackets = ')]}';
      MaxDictVer = 7;                       {highest dictionary version}
var HDict: THandle;
    st, gl, sel: string;
    i, j, k, l, m, n, p, ver: Integer;
    h: Byte;
    br: Boolean;
    r: DWord;
    t: TDateTime;
    ach: array[0..255] of Byte;
    lach: array[0..255] of Integer;
begin
  t:=CompileTime(Options.FileListBox3.FileName);
  if t<>DictWriteTime then begin
    DictWriteTime:=t;
    EntriesLabel.Font.Color:=clRed;
    with Dictionary do if ItemIndex=-1 then sel:='' else sel:=Items.Names[Dictionary.ItemIndex];
    Dictionary.Items.Clear;
    Dictionary.Items.BeginUpdate;
    if not FileExists(FileListBox3.FileName) then EntriesLabel.Caption:='No existing dictionary file (*.ngr) selected' else begin
      HDict:=CreateFile(PChar(FileListBox3.FileName), Generic_Read, File_Share_Read or File_Share_Write, nil, Open_Always, File_Attribute_Normal, 0);
      if HDict=Invalid_Handle_Value then EntriesLabel.Caption:='Error opening dictionary' else begin
        ReadFile(HDict, l,   4,  r, nil);
        ReadFile(HDict, ach, l,  r, nil); {"NGR" + version id + selected language}
        ver:=ach[3]-32;
        if ver<=MaxDictVer then begin
          ReadFile(HDict, n,   4,  r, nil); {selected word}
          ReadFile(HDict, n,   4,  r, nil); {number of languages}
          if r=0 then n:=0;
          ReadFile(HDict, lach,  4*n,   r, nil);
          Dictionary.Items.Capacity:=lach[lModLem];
          for i:=0 to Min(n-1, lModLem) do for j:=0 to lach[i]-1 do begin
            ReadFile(HDict, k,     4,   r, nil);
            ReadFile(HDict, ach,   k,   r, nil);
            st:='';
            if i=lModLem then for l:=0 to k-1 do st:=st+Letters[ach[l]+1];
            ReadFile(HDict, ach,   k,   r, nil);
            if i=lModLem then for l:=0 to k-1 do if ach[l]=2 then st[l+1]:='‡';
            ReadFile(HDict, k,     2,   r, nil);  {ending}
            if ver>=3 then ReadFile(HDict, k, 2, r, nil);  {morpheme boundary}
            ReadFile(HDict, k,     4,   r, nil);  {gloss}
            ReadFile(HDict, ach,   k,   r, nil);
            ach[r]:=0;
            gl:=PChar(@ach);
            br:=False;
            for l:=1 to 3 do begin
              p:=Pos(openbrackets[l], gl);
              if p>0 then Delete(gl, p, Pos(closebrackets[l], gl)-p+1);
              br:=br or (p>0);
            end;
            ReadFile(HDict, h,     1,   r, nil);  {confirmed}
            p:=Dictionary.Items.IndexOfName(Trim(gl));
            if (i=lModLem) and (UnconfirmedBox.Checked or (h=1)) and (Trim(gl)<>'') then begin
              if (p>-1) and not br then Dictionary.Items.Delete(p);
              if (p=-1) or  not br then Dictionary.Items.AddObject(Trim(Copy(gl, 1, MultPos(',', gl)-1))+'='+st, Pointer(h));
            end;
            ReadFile(HDict, m,     4,   r, nil);  {connotations}
            for k:=0 to m-1 do begin
              ReadFile(HDict, h,   1,   r, nil);
              if ver=4 then ReadFile(HDict, h,   1,   r, nil);
              ReadFile(HDict, l,   4,   r, nil);
              ReadFile(HDict, ach, l,   r, nil);
            end;
            ReadFile(HDict, h,     1,   r, nil);  {ancestor-lang}
            if (ver<=2) or (h<255)           then ReadFile(HDict, ach,   4, r, nil);  {ancestor}
            if (ver=2) or (ver>=3) and (h=i) then ReadFile(HDict, l,     4, r, nil);  {compound part 2}
            if ver>=1                        then ReadFile(HDict, h,     1, r, nil);  {variant code}
            ReadFile(HDict, k,     4,   r, nil);  {irregular}
            ReadFile(HDict, ach,   k,   r, nil);
            ReadFile(HDict, k,     4,   r, nil);  {descendants}
            ReadFile(HDict, ach,   5*k, r, nil);
            if ver>=6                        then ReadFile(HDict, h,     1,   r, nil);  {mark}
          end;
        end;
        if Dictionary.Items.Count>0 then EntriesLabel.ParentFont:=True;
        EntriesLabel.Caption:=IntToStr(Dictionary.Items.Count)+' lemma'+IfThen(Dictionary.Items.Count<>1, 'ta')+' in dictionary';
        NgrVerLabel.Caption:='Thatís a version '+IntToStr(ver)+' dictionary file'+IfThen(ver>MaxDictVer, ', too high to read')+'.';
      end;
      CloseHandle(HDict);
    end;
    with Dictionary do ItemIndex:=Items.IndexOfName(sel);
    Dictionary.Items.EndUpdate;
  end;
end; {UpdateDict}

procedure TOptions.LemBoxClick(Sender: TObject);
begin
  Dictionary.Invalidate;
end;

procedure TOptions.DictionaryDrawItem(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
var p: Integer;
begin
  with Dictionary.Canvas, Dictionary.Items do begin
    FillRect(Rect);
    if Integer(Objects[Index])=0 then Font.Color:=clSilver;
    Font.Name:=Serif;
    TextOut(Rect.Left+2, Rect.Top, Names[Index]);
    if LemBox.Checked then Font.Name:=LemFont else Font.Name:=SansSerif;
    p:=Pos('=', Strings[Index]);
    TextOut(Rect.Left+Dictionary.Width div 2, Rect.Top, Copy(Strings[Index], p+1, 255)+'.');
  end;
end; {DictionaryDrawItem}

procedure TOptions.StatListBoxDrawItem(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
var p, q: Integer;
    st, st2: string;
begin
  StatListBox.Canvas.FillRect(Rect);
  st:=StatListBox.Items[Index];
  with StatListBox.Canvas do if (Index>=32) and (Index<36) then TextOut(Rect.Left, Rect.Top, st) else begin
    if Index<32 then begin
      Font.Name:=LemFont;
      TextOut(Rect.Left+3, Rect.Top, st[1]);
      Font.Name:=SansSerif;
      TextOut(Rect.Left+15, Rect.Top, ':');
    end else;
    p:=Pos(':', st);
    q:=Pos('_', st);
    st2:=Copy(st, p+1, q-p-1);
    TextOut(Rect.Left +60-TextWidth(st2), Rect.Top, st2);
    st2:=Copy(st, q+1, MaxInt);
    TextOut(Rect.Left+120-TextWidth(st2), Rect.Top, st2);
    if Index>=36 then TextOut(Rect.Left+130, Rect.Top, Copy(st, 1, p-1));
  end;
end; {StatListBoxDrawItem}

procedure TOptions.NeogrammBtnClick(Sender: TObject);
begin
  TableMaker.NeogrammBtnClick(nil);
end; {NeogrammBtnClick}

function SubstituteStems(S: string; Lemizh: Boolean; Dict: TStrings): string;
const ending: array[0..4] of string = ('aeyioOuU‡ËÃÏÚ“˘Ÿ·È˝ÌÛ”˙⁄', 'rRlnm', 'nm', '', '');
      brackets = '([{}])';
      bracketset  = ['{', '}', #21..Char(20+Length(brackets))];
      obracketset = ['{',      #21..Char(20+Length(brackets) div 2)];
var p, q, r, u, a, i: Integer;
    lem, c, innercase, poststem, st: string;
begin
  if TableMaker.SubstStems then begin
    Result:='';
    if Lemizh then begin
      for i:=1 to Length(brackets) do S:=StringReplace(S, '\'+brackets[i], Char(20+i), [rfReplaceAll]);
      repeat
        p:=Pos('(', S);
        if p>0 then begin
          q:=Pos(')', Copy(S, p+1, MaxInt));
          if q=0 then q:=Length(S);
          r:=1;  i:=0;
          repeat
            c:=Copy(S, p+q+r, 1);
            if Pos(c, ending[i])>0 then begin
              Inc(i);
              if Pos(c, ending[i])>0 then Inc(i);
            end else if (c='') or not (c[1] in bracketset) then Break;
            Inc(r);
          until i=4;
          repeat
            c:=Copy(S, p+q+r-1, 1);
            if (c='') or (c[1] in obracketset) then Dec(r);
          until (c='') or not (c[1] in obracketset);
          st:=Copy(S, p+1, q-1);
          u:=Dict.IndexOfName(st);
          if (u>-1) and (Dict.Names[u]<>st) then TangoMessageBox('Capitalisation of stem ë'+st+'í does not match dictionary.', mtWarning, [mbOK], '');
          lem:=Dict.Values[st];
          if lem<>'' then begin
            innercase:=Copy(S, p+q+1, r-1);
            a:=Pos('‡', lem);
            poststem:=Copy(lem, a+1, MaxInt);
            if Copy(innercase, MultPos('{}', innercase), 1)='}' then poststem:='{'+poststem+'}';
            if (poststem<>'') and (Copy(innercase, LastDelimiter('{}', innercase), 1)='{') then poststem:='}'+poststem+'{';
            Result:=Result+Copy(S, 1, p-1){preceding chars}+Copy(lem, 1, a-1){prestem}+innercase+poststem;
          end else Result:=Result+Copy(S, 1, p+q+r-1);
          S:=Copy(S, p+q+r, MaxInt);
        end;
      until p=0;
      Result:=Result+S;
      for i:=1 to Length(brackets) do Result:=StringReplace(Result, Char(20+i), '\'+brackets[i], [rfReplaceAll]);
    end else repeat
      p:=Pos('<lem>', S);
      if p=0 then Result:=Result+S else begin
        q:=Pos('</lem>', S);
        Result:=Result+Copy(S, 1, p+4)+SubstituteStems(Copy(S, p+5, q-p-5), True, Dict)+'</lem>';
        S:=Copy(S, q+6, MaxInt);
      end;
    until p=0;
  end else Result:=S;
  Result:=StringReplace(StringReplace(Result, '{}', '', [rfReplaceAll]), '}{', '', [rfReplaceAll]);
end; {SubstituteStems}

{------------------------------------------------ Unglossed & Statistics -----------------------------------------------------------}

procedure TOptions.PageControlChange(Sender: TObject);
const letters = 'aeyioOuUlRrnmgdbktpjczvwxhsqf,.-‡ËÃÏÚ“˘Ÿ·È˝ÌÛ”˙⁄';
var sl: TStringList;
    al: array[0..Length(letters)] of Integer;
  procedure FillSL(S: string);
  var i, p, q, r, n: Integer;
      lemma: string;
  begin
    if Copy(S, 1, 4)='<id=' then S:=TrimLeft(Copy(S, Pos('>', S)+1, MaxInt));
    S:=StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(S, '\(', '', [rfReplaceAll]), '\(', '', [rfReplaceAll]),
      '<nut>', '', [rfReplaceAll]), '</nut>', '', [rfReplaceAll]), '<nonut>', '', [rfReplaceAll]), '</nonut>', '', [rfReplaceAll]);
    if (S<>'') and ((S[1]<>'(') or (S[Length(S)]<>')')) then repeat
      p:=Pos('(', S);
      for i:=1 to IfThen(p>0, p-1, Length(S)) do Inc(al[Pos(S[i], letters)]);
      if p>0 then begin
        q:=MultPos('()\'#10#13, Copy(S, p+1, MaxInt));
        if (Copy(S, p-1, 1)<>'\') and (Copy(S, p+q, 1)=')') then begin
          lemma:=Copy(S, p+1, q-1);
          r:=sl.IndexOf(lemma);
          if r=-1 then n:=1 else n:=Integer(sl.Objects[r])+1;
          sl.AddObject(lemma, Pointer(n));
        end;
        S:=Copy(S, p+q+1, MaxInt);
      end;
    until p=0;
  end;
var i, j, k, l, p, q: Integer;
    st: string;
    sl2: TStringList;
begin
  if Sender=UnglossedInclText then StatInclText.Checked:=UnglossedInclText.Checked else if Sender=StatInclText then UnglossedInclText.Checked:=StatInclText.Checked;
  FindPanel.Visible:=(PageControl.ActivePage=DictSheet) or (PageControl.ActivePage=GlossedSheet);
  if (PageControl.ActivePage=GlossedSheet) or (PageControl.ActivePage=StatSheet) then with TableMaker do begin
    for i:=0 to Length(letters) do al[i]:=0;
    sl:=TStringList.Create;
    sl.Sorted:=True;
    sl.Duplicates:=dupIgnore;
    for i:=0 to Length(Heads)-2 do for j:=0 to Length(Tables[i])-1 do for k:=0 to Length(Tables[i,j].HTMLs)-1 do
        for l:=0 to ColCounts[i]-1+Ord(i=SolvTab) do if ColType[i,l]=ctLem then FillSL(Tables[i,j].HTMLs[k,l]) else if ColType[i,l]<>ctSolv then begin
      st:=Tables[i,j].HTMLs[k,l];
      repeat
        p:=Pos('<lem>', st);
        if p>0 then begin
          q:=Pos('</lem>', st);
          FillSL(Copy(st, p+5, q-p-5));
          st:=Copy(st, q+6, MaxInt);
        end;
      until p=0;
    end;
    for i:=1-Ord(UnglossedInclText.Checked) to Length(TextPages)-1 do FillSL(TextPages[i].Text+' ');
    sl2:=TStringList.Create;
    for i:=0 to Dictionary.Items.Count-1 do begin
      j:=sl.IndexOf(Dictionary.Items.Names[i]);
      if j>-1 then begin
        sl2.Add(Dictionary.Items.Names[i]);
        st:=Copy(Dictionary.Items[i], Pos('=', Dictionary.Items[i])+1, MaxInt);
        for k:=1 to Length(st) do if st[k]<>'‡' then Inc(al[Pos(st[k], letters)], Integer(sl.Objects[j]));
        sl.Delete(j);
      end;
    end;
    if not (sl.Equals(UnglossedMemo.Lines) and sl2.Equals(GlossedMemo.Lines)) then begin
      UnglossedMemo.Lines.Assign(sl);
      GlossedMemo.Lines.Assign(sl2);
      sl.Free;
      sl2.Free;
      UnglossedMemo.SelStart:=0;     GlossedMemo.SelStart:=0;
      i:=UnglossedMemo.Lines.Count;  j:=GlossedMemo.Lines.Count;
      UnglossedLabel.Visible:=i+j>0;
      if UnglossedLabel.Visible then begin
        UnglossedLabel.Caption:=IntToStr(i)+' ('+FormatFloat('#0.0', 100*i/(i+j+0.0))+'%) unglossed word'+IfThen(i<>1, 's')+
                        ' and '+IntToStr(j)+' ('+FormatFloat('#0.0', 100*j/(i+j+0.0))+'%) glossed word'+IfThen(j<>1, 's')+' found of '+
                        IntToStr(i+j);
        UnglossedGauge.Progress:=Round(100*j/(i+j));
      end;
      if FindEdit.Text<>'' then FindDownClick(nil);
    end;
    StatListBox.Clear;
    j:=0;
    for i:=1 to Length(letters) do Inc(j, al[i]);
    for i:=1 to 8 do al[i]:=al[i]+al[i+29]+al[i+37];
    for i:=1 to Length(letters)-16 do StatListBox.Items.Add(letters[i]+':'+IntToStr(al[i])+'_('+FormatFloat('#0.0', 100*al[i]/j)+'%)');
    StatListBox.Items.Add('');
    StatListBox.Items.Add(IntToStr(j)+' letters and basic punctuation altogether');
    if UnglossedLabel.Visible and (UnglossedGauge.Progress<95) and (UnglossedGauge.Progress>0) then
      StatListBox.Items.Add('Vowels are overrepresented by a factor of ~'+FormatFloat('###0.0', 100/(UnglossedGauge.Progress))+'.') else
        StatListBox.Items.Add('');
    for i:=2 to 7 do begin
      al[1]:=al[1]+al[i];  al[30]:=al[30]+al[29+i];  al[38]:=al[38]+al[37+i];
    end;
    al[39]:=al[1]-al[30]-al[38];
    if al[1]>0 then begin
      StatListBox.Items.Add('');
      StatListBox.Items.Add('vowels with low accent:' +IntToStr(al[30])+'_('+FormatFloat('#0.0', 100*al[30]/al[1])+'%)');
      StatListBox.Items.Add('vowels with high accent:'+IntToStr(al[38])+'_('+FormatFloat('#0.0', 100*al[38]/al[1])+'%)');
      StatListBox.Items.Add('vowels without accent:'  +IntToStr(al[39])+'_('+FormatFloat('#0.0', 100*al[39]/al[1])+'%)');
    end;
  end;
end; {PageControlChange}

procedure TOptions.UnglossedLabelMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
begin
  UnglossedPanel.Show;
  UnglossedTimer.Enabled:=True;
end; {UnglossedLabelMouseMove}

procedure TOptions.UnglossedTimerTimer(Sender: TObject);
var p1, p2: TPoint;
begin
  p1:=UnglossedPanel.ScreenToClient(Mouse.CursorPos);
  p2:=UnglossedLabel.ScreenToClient(Mouse.CursorPos);
  if not (((p1.X>=0) and (p1.X<=UnglossedPanel.Width) and (p1.Y>=0) and (p1.Y<=UnglossedPanel.Height))
       or ((p2.X>=0) and (p2.X<=UnglossedLabel.Width) and (p2.Y>=0) and (P2.Y<=UnglossedLabel.Height))) then begin
    UnglossedPanel.Hide;
    UnglossedTimer.Enabled:=False;
  end;
end; {UnglossedTimerTimer}

{------------------------------------------------ Find -----------------------------------------------------------}

procedure TOptions.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key=VK_F3) and (Shift=[]) then begin
    FocusControl(FindEdit);
    FindEdit.SelectAll;
    if FindClear.Enabled then FindDownClick(nil);
  end;
end; {FormKeyDown}

procedure TOptions.FindClearClick(Sender: TObject);
begin
  FindEdit.Text:='';
  FindEdit.Color:=clWindow;
  FindEdit.Tag:=0;
end; {FindClearClick}

procedure TOptions.FindEditChange(Sender: TObject);
const clSuccess: array[False..True] of TColor = ($c0c0ff, clWindow);
var p: Integer;
    f: Boolean;
    g1, g2: TMemo;
begin
  FindClear.Enabled:=FindEdit.Text<>'';
  FindDown.Enabled:=FindClear.Enabled;
  FindEdit.ParentFont:=True;
  if PageControl.ActivePage=DictSheet then begin
    with Dictionary do if (Items.Count>0) and (ItemIndex=-1) then ItemIndex:=0;
    f:=False;
    p:=Dictionary.ItemIndex-1;
    with Dictionary.Items do if Count>0 then repeat
      p:=(p+1) mod Count;
      f:=(Pos(LowerCase(FindEdit.Text), LowerCase(Names[p]))>0) or (Pos(RemoveDiacritics(FindEdit.Text), RemoveDiacritics(Values[Names[p]])+'.')>0);
    until f or (p=(Dictionary.ItemIndex+Count-1) mod Count);
    FindEdit.Color:=clSuccess[f or not FindClear.Enabled];
    if f then Dictionary.ItemIndex:=p;
  end else begin {GlossedSheet}
    if FindEdit.Tag=0 then begin
      g1:=UnglossedMemo;  g2:=GlossedMemo;
    end else begin
      g1:=GlossedMemo;    g2:=UnglossedMemo;
    end;
    g1.SelLength:=0;
    g2.SelLength:=0;
    p:=Pos(LowerCase(FindEdit.Text), LowerCase(Copy(g1.Lines.Text, g1.SelStart+1, MaxInt)+g2.Lines.Text+Copy(g1.Lines.Text, 1, g1.SelStart)));
    FindEdit.Color:=clSuccess[(p>0) or not FindClear.Enabled];
    if p>0 then begin
      p:=(p+g1.SelStart) mod (Length(g1.Lines.Text)+Length(g2.Lines.Text));
      FindEdit.Tag:=FindEdit.Tag xor Ord(p>Length(g1.Lines.Text));
      if p<=Length(g1.Lines.Text) then begin
        g1.SelStart:=p-1;
        g1.SelLength:=Length(FindEdit.Text);
      end else begin
        g2.SelStart:=p-1-Length(g1.Lines.Text);
        g2.SelLength:=Length(FindEdit.Text);
      end;
    end;
  end;
end; {FindEditChange}

procedure TOptions.FindDownClick(Sender: TObject);
begin
  if PageControl.ActivePage=DictSheet then
    with Dictionary do ItemIndex:=(ItemIndex+1) mod Items.Count
  else begin {GlossedSheet}
    if FindEdit.Tag=0 then with UnglossedMemo do SelStart:=SelStart+SelLength+1 else with GlossedMemo do SelStart:=SelStart+SelLength+1;
  end;
  FindEditChange(nil);
end; {FindDownClick}

end.

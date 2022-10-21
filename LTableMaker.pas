unit LTableMaker;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, Math, FileCtrl,
  StdCtrls, Buttons, Grids, ComCtrls, ExtCtrls, CLipbrd, ToolWin, MyUtils, Registry,
  ImgList, Spin, Menus, Random;

const
  ctEng=0;  ctLit=1; ctLem=2; ctSolv=3; ctHead=4;
  Capt: array[ctEng..ctHead] of string = ('English', 'Literal', 'Lemizh', 'Solution no.', 'Header');
  Heads: array[0..3] of string = ('Examples', 'Exercises', 'Vocabulary', 'Text');
  Text0 = 'sandbox ';
  ColCounts: array[0..Length(Heads)-2] of Integer = (4, 3, 2);
  ColType: array[0..Length(Heads)-2, 0..3] of Integer = ((ctHead, ctLem, ctLit, ctEng), (ctHead, ctLem, ctSolv, 0), (ctLem, ctEng, 0, 0));
  Serif = 'Gentium';   SansSerif = 'Arial';   LemFont = 'Lemizh';
  clActive: array[False..True] of TColor = (clGrayText, clWindowText);
  clSuccess: array[False..True, False..True] of TColor = (($c0c0ff, $c0e8ff), (clWindow, clWindow));
  NutTags: array[0..1, False..True] of string = (('<nonut>', '<nut>'), ('</nonut>', '</nut>'));
  NutColors: array[False..True, False..True] of TColor = (($dddddd, $ff77bb), (clGray, $77bbff));
  BottomShortcuts = #13#10'F9 (+ Shift / Ctrl / Alt / Shift+Ctrl / Shift+Alt): Options'#13#10'F10: Open Neogrammarian'#13#10#13#10;
  AllShortcuts = 'F2: Editor'#13#10'F3: Find'#13#10'F4: Select multiple lines'#13#10'F5: Suppress line(s) for nutshells'+BottomShortcuts+'Ctrl+1/2/3: <lem, nut, nonut> tags'#13#10'Ctrl+7/8/9: Brackets';
  MemoShortcuts = 'F3: Find'+BottomShortcuts+'Ctrl+1: <lem> tags'#13#10'Ctrl+7/8/9: Brackets';
  MemoSyntax = '"\": Escape bracket or case gloss'#13#10'"|": OR between alternative case endings'#13#10'"{´}": Highlight (non)agent';
  TableSyntax = '">" in Header column: Separate from previous line'#13#10'"<id=...>": HTML id'#13#10'[space] at Lemizh cell start: Suppress gloss'#13#10'"xxx": Mark as to do'#13#10+MemoSyntax;
  SolvSyntax = '[space] at line start: Suppress gloss'#13#10+MemoSyntax;
  MarksHi = 6;
  SymLetters = 'con';
  MaxAppendices = 16; {total number of "units" must not exceed 100}

type
  TTable = class(TObject)
    Un, Nut: Byte;
    Id: string;
    Mark: Byte;
    HTMLs: array of array[0..3] of string;
    NoNutshell: array of Boolean;
  end;
  
  TTextPage = record
    Text, Header: string;
    ColSpan, RowSpan, Caret, Selection: Integer;
  end;

  TTableMaker = class(TForm)
    TabControl: TTabControl;     ImageList: TImageList;       BottomPanel: TPanel;         LemBar: TToolBar;            ToolButton1: TToolButton;    ToolButton2: TToolButton;    ToolButton3: TToolButton;
    ToolButton4: TToolButton;    ToolButton5: TToolButton;    ToolButton6: TToolButton;    ToolButton7: TToolButton;    ToolButton8: TToolButton;    ToolButton9: TToolButton;    ToolButton10: TToolButton;
    ToolButton11: TToolButton;   ToolButton12: TToolButton;   ToolButton13: TToolButton;   ToolButton14: TToolButton;   ToolButton15: TToolButton;   ToolButton16: TToolButton;   ToolButton17: TToolButton;
    ToolButton18: TToolButton;   ToolButton19: TToolButton;   ToolButton20: TToolButton;   ToolButton21: TToolButton;   ToolButton22: TToolButton;   ToolButton23: TToolButton;   ToolButton24: TToolButton;
    ToolButton25: TToolButton;   ToolButton26: TToolButton;   ToolButton27: TToolButton;   ToolButton28: TToolButton;   ToolButton29: TToolButton;   ToolButton30: TToolButton;   ToolButton31: TToolButton;
    ToolButton32: TToolButton;   ToolButton33: TToolButton;   ToolButton34: TToolButton;   ToolButton35: TToolButton;   ToolButton36: TToolButton;   ToolButton37: TToolButton;   ToolButton38: TToolButton;
    ToolButton39: TToolButton;   ToolButton40: TToolButton;   ToolButton41: TToolButton;   ToolButton42: TToolButton;   ToolButton43: TToolButton;   ToolButton44: TToolButton;   ToolButton45: TToolButton;
    OKBtn: TBitBtn;              CancelBtn: TBitBtn;          SaveBtn: TBitBtn;            Image1: TImage;              ToolLevel: TSpeedButton;     ToolHilight: TSpeedButton;   ToolStem: TSpeedButton;
    LicenseImage: TImage;        VLabel: TLabel;              Notebook: TNotebook;         Memo: TMemo;                 CutMemoBtn: TBitBtn;         CopyMemoBtn: TBitBtn;        DelMemoBtn: TBitBtn;
    ListPanel: TPanel;           ListBox: TListBox;           StaticText3: TStaticText;    StaticText4: TStaticText;    NewTableBtn: TBitBtn;        DelTableBtn: TBitBtn;        UpTableBtn: TBitBtn;
    DownTableBtn: TBitBtn;       GridPanel: TPanel;           StringGrid: TStringGrid;     UnitCombo: TComboBox;        IdEdit: TEdit;               CopyBtn: TBitBtn;            NewLineBtn: TBitBtn;
    StaticText1: TStaticText;    StaticText2: TStaticText;    UpLineBtn: TBitBtn;          DownLineBtn: TBitBtn;        Splitter: TSplitter;         TextTagsCombo: TComboBox;    Label1: TLabel;
    Label2: TLabel;              LemBox: TCheckBox;           SpanPanel: TPanel;           Label3: TLabel;              ColspanEdit: TSpinEdit;      RowspanEdit: TSpinEdit;      Label4: TLabel;
    ToolButton48: TToolButton;   ToolButton49: TToolButton;   FindPanel: TPanel;           FindEdit: TEdit;             Label5: TLabel;              FindUp: TSpeedButton;        CaseOptsBtn: TSpeedButton;
    FindDown: TSpeedButton;      ToolButton50: TToolButton;   ToolButton51: TToolButton;   ToolButton52: TToolButton;   FindClear: TSpeedButton;     CLabel: TLabel;              LemizhImage: TImage;
    CaseMenu: TPopupMenu;        CaseBtn: TSpeedButton;       ToolButton55: TToolButton;   ToolButtonAstro: TToolButton;ToolButton57: TToolButton;   ToolButton58: TToolButton;   RandomBtn: TBitBtn;
    ToolButton59: TToolButton;   ToolButton60: TToolButton;   ToolButton61: TToolButton;   ToolButton62: TToolButton;   ToolButton63: TToolButton;   ClipboardMenu: TPopupMenu;   Deletelines1: TMenuItem;
    Cutlines1: TMenuItem;        Copylines1: TMenuItem;       Pastelines1: TMenuItem;      LinesBtn: TBitBtn;           MultiSelectBtn: TSpeedButton;ToolLem: TSpeedButton;       ToolButton53: TToolButton;
    ToolButton64: TToolButton;   ToolButton46: TToolButton;   ToolButton47: TToolButton;   ToolButton65: TToolButton;   ToolButton66: TToolButton;   ToolButton67: TToolButton;   SpeedButton2: TSpeedButton;
    ToolButton68: TToolButton;   NrTables: TStaticText;       Image3: TImage;              FindPanel2: TPanel;          AccentCheckBox: TCheckBox;   LemFindBtn: TSpeedButton;    VowelCheckBox: TCheckBox;
    CaseCheckBox: TCheckBox;     LinesPopup: TPopupMenu;      Add1: TMenuItem;             Delete1: TMenuItem;          N1: TMenuItem;               Up1: TMenuItem;              Down1: TMenuItem;
    Addline1: TMenuItem;         Cutline1: TMenuItem;         Pasteline1: TMenuItem;       Copyline1: TMenuItem;        Deleteline1: TMenuItem;      N2: TMenuItem;               Up2: TMenuItem;
    Down2: TMenuItem;            N3: TMenuItem;               TablePopup: TPopupMenu;      N4: TMenuItem;               Mark1: TMenuItem;            Copytable1: TMenuItem;       Sortsolutions1: TMenuItem;
    Edit1: TMenuItem;            SaveCaretBox: TCheckBox;     CopyMemoGlossBtn: TBitBtn;   ToolButton69: TToolButton;   ToolButton70: TToolButton;   ToolButton71: TToolButton;   ToolButton72: TToolButton;
    TimeLabel: TLabel;           N6: TMenuItem;               ToolButton74: TToolButton;   NeogrammBtn: TBitBtn;        MemoTimeLabel: TLabel;       StaticText5: TStaticText;    Selectmlines1: TMenuItem;
    UnitList: TListBox;          SolvPanel: TPanel;           SolvMemo: TMemo;             SolvOKBtn: TBitBtn;          SolvCaption: TLabel;         Editsolution1: TMenuItem;    N5: TMenuItem;
    NutshellBtn: TSpeedButton;   CopyNutshellBtn: TBitBtn;    ShortcutsImage: TImage;      MemoShortcutsImage: TImage;  Selectall1: TMenuItem;       N7: TMenuItem;               Copytablefornutshell1: TMenuItem;
    Selectall2: TMenuItem;       N8: TMenuItem;               NutPanel: TPanel;            ToolNut: TSpeedButton;       ToolNonut: TSpeedButton;     CopyMemoBothBtn: TBitBtn;    Suppressfornutshell1: TMenuItem;
    StaticText6: TStaticText;    NutCombo: TComboBox;         Mark2: TMenuItem;            Mark3: TMenuItem;            Mark6: TMenuItem;            Mark4: TMenuItem;            MemoSyntaxLabel: TLabel;
    Mark5: TMenuItem;            HTML5Panel: TPanel;          HTML5Image: TImage;          SyntaxLabel: TLabel;         AstroMenu: TPopupMenu;       Sun1: TMenuItem;             Moon1: TMenuItem;
    Mercury1: TMenuItem;         Venus1: TMenuItem;           Earth1: TMenuItem;           Mars1: TMenuItem;            Jupiter1: TMenuItem;         Saturn1: TMenuItem;          recurringpart1: TMenuItem;
    Uranus1: TMenuItem;          Neptune1: TMenuItem;         negation1: TMenuItem;        N9: TMenuItem;               waningMoon1: TMenuItem;      waningMercury1: TMenuItem;   waningVenus1: TMenuItem;
    n10: TMenuItem;              hl1: TMenuItem;              hr1: TMenuItem;              h1: TMenuItem;               N11: TMenuItem;              n12: TMenuItem;              m1: TMenuItem;
    N13: TMenuItem;              N14: TMenuItem;              minus1: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure TabControlChange(Sender: TObject);
    procedure ListBoxClick(Sender: TObject);
    procedure NewTableBtnClick(Sender: TObject);
    procedure DelTableBtnClick(Sender: TObject);
    procedure CopyBtnClick(Sender: TObject);
    procedure ListBoxDrawItem(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
    procedure UnitOrIdChange(Sender: TObject);
    procedure StringGridSetEditText(Sender: TObject; ACol, ARow: Integer; const Value: string);
    procedure NewLineBtnClick(Sender: TObject);
    procedure Deletelines1Click(Sender: TObject);
    procedure StringGridDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect; State: TGridDrawState);
    procedure StringGridGetEditText(Sender: TObject; ACol, ARow: Integer; var Value: string);
    procedure LemButtonClick(Sender: TObject);
    procedure ControlEnterExit(Sender: TObject);
    procedure UpDownTableBtnClick(Sender: TObject);
    procedure UpDownLineBtnClick(Sender: TObject);
    procedure StringGridKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure DelMemoBtnClick(Sender: TObject);
    procedure CopyMemoBtnClick(Sender: TObject);
    procedure CutMemoBtnClick(Sender: TObject);
    procedure MemoChange(Sender: TObject);
    procedure MemoKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure TabControlChanging(Sender: TObject; var AllowChange: Boolean);
    procedure OKBtnClick(Sender: TObject);
    procedure CancelBtnClick(Sender: TObject);
    procedure StringGridClick(Sender: TObject);
    procedure SaveBtnClick(Sender: TObject);
    procedure ToolHilightClick(Sender: TObject);
    procedure LemBoxClick(Sender: TObject);
    procedure LinkImageClick(Sender: TObject);
    procedure AppShowHint(var HintStr: string; var CanShow: Boolean; var HintInfo: THintInfo);
    procedure TextTagsComboChange(Sender: TObject);
    procedure StringGridDblClick(Sender: TObject);
    procedure FindEditChange(Sender: TObject);
    procedure FindClearClick(Sender: TObject);
    procedure CaseBtnClick(Sender: TObject);
    procedure SymbolClick(Sender: TObject);
    procedure StringGridSelectCell(Sender: TObject; ACol, ARow: Integer; var CanSelect: Boolean);
    procedure LinesBtnClick(Sender: TObject);
    procedure Cutlines1Click(Sender: TObject);
    procedure Copylines1Click(Sender: TObject);
    procedure Pastelines1Click(Sender: TObject);
    procedure MultiSelectBtnClick(Sender: TObject);
    procedure OptsBtnClick(Sender: TObject);
    procedure ListBoxDblClick(Sender: TObject);
    procedure FindEditKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure LemFindBtnClick(Sender: TObject);
    procedure SplitterMoved(Sender: TObject);
    procedure FormMouseWheel(Sender: TObject; Shift: TShiftState; WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
    procedure PopupPopup(Sender: TObject);
    procedure Edit1Click(Sender: TObject);
    procedure StringGridMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure FormShow(Sender: TObject);
    procedure NeogrammBtnClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UnitListChange(Sender: TObject);
    procedure MemoMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure SolvOKBtnClick(Sender: TObject);
    procedure SolvMemoChange(Sender: TObject);
    procedure Editsolution1Click(Sender: TObject);
    procedure Sortsolutions1Click(Sender: TObject);
    procedure NutshellBtnClick(Sender: TObject);
    procedure Selectall1Click(Sender: TObject);
    procedure NutComboChange(Sender: TObject);
    procedure MemoKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure HTML5ImageDblClick(Sender: TObject);
    procedure RandomBtnClick(Sender: TObject);
    procedure AstroMenuClick(Sender: TObject);
    procedure ToolButtonAstroClick(Sender: TObject);
  private
    ShallSave, MouseC, MouseR, SolvX, SolvY, LemFontCorr: Integer;
    CF_Rows: array[0..Length(Heads)-2] of Word;
    BackupFiles, InterGlossing, UnconfirmedStems, OptsChanged, CaseMenuBreak, FoundTitle, UploadData, UploadSolv, ShowUploadWindow: Boolean;
    MemoFontSize: Integer;
    SavePath, SolutionsPath, OldSolutionsPath, DictPath, DictFile, NeogrammPath, HostName, UserName, Password, UploadPath, UploadSolvPath: string;
    ListBoxIndex: array[0..Length(Heads)-2] of Integer;
    ColWidths: array[0..Length(Heads)-2, 0..3] of Integer;
    Backups, MarkBackups: array[0..Length(Heads)-1] of string;
    procedure AppActivate(Sender: TObject);
    procedure SetLemFont(Font: TFont; CanLem: Boolean; SizeOffset: Integer);
    procedure UpdateTableButtons;
    procedure UpdateCasePopup;
    procedure UpdateTextPages;
    procedure MakeFileAndMarks(Nr: Integer; out AFile, AMark: string);
    function SaveWhat(out Files, Marks: array of string): string;
    procedure Save(All: Boolean; Files, Marks: array of string);
    function NewLine(Strings: array of string; Line: Integer): Integer;
    procedure EnableLemBar(Col: Integer);
    procedure ExchangeTables(I, J: Integer);
    procedure ListBoxModified(Focus: Integer);
    procedure SetCaption;
    procedure CaseClick(Sender: TObject);
    function IncludeSymbols(S: string): string;
    function IncludeBracketsEtc(S: string; Symbols: Boolean): string;
    function TestLemParts(S: string; Lemizh: Boolean; var ErrMsg: string): string;
    function Glossing(S: string; var ErrMsg: string): string;
    function IncludeAbbrEtc(S: string; var ErrMsg: string): string;
    function SelCaretLine: Integer;
    procedure ShowOpts(TabNr: Integer);
    function GetFreeSolutionNo(ACol, ARow: Integer): string;
    function FullSolutionsPath: string;
    function GetNoOfSolutions(A, B, Un: Integer; Id: string; out No, SolvCol: Integer): string;
    function ThisTable: TTable;
    function HTMLSpans: string;
    function HTML4or5(S: string): string;
    procedure ErrorMessageBox(Err: string);
  public
    SubstStems: Boolean;
    SolvTab: Integer;
    OptsFind: string;
    CaseAbbrs, CaseExtds: array[0..33] of string;
    Symbols: TTranscSymbols;
    Tables: array[0..Length(Heads)-2] of array of TTable;
    TextPages: array of TTextPage;
    function SaveSolvHTMLs(Always: Boolean): string;
  end;

var
  TableMaker: TTableMaker;

function CaseEnding(i: Integer): string;

implementation

uses Opts, SaveSol;

{$R *.DFM}

procedure TTableMaker.FormCreate(Sender: TObject);
  procedure SplitIniLine(S: string; AComboBox: TComboBox);
  var p: Integer;
  begin
    repeat
      p:=Pos('|', S);
      AComboBox.Items.Add(StringReplace(Copy(S, 1, p-1), '\u124?', '|', [rfReplaceAll]));
      S:=Copy(S, p+1, MaxInt);
    until S='';
  end; {SplitIniLine}
var i, j, k, p, q, q2: Integer;
    sl: TStringList;
    Reg: TRegIniFile;
    st: string;
    tm, dt: TDateTime;
    ach: array[0..1024] of Char;
begin
  tm:=Now;
  DecimalSeparator:='.';
  ThousandSeparator:=',';
  Application.OnShowHint:=AppShowHint;
  SolvMemo.Font.Name:=Serif;
 // MemoSyntaxLabel.Hint:=MemoSyntax;
  MemoShortcutsImage.Hint:=MemoShortcuts;
  ShallSave:=mrCancel;
  Reg:=TRegIniFile.Create('SOFTWARE\TESSoft Inc.\LTable');
  BackupFiles:=Reg.ReadBool('Save', 'Backup', True);
  GetEnvironmentVariable('APPDATA', ach, SizeOf(ach));
  SavePath:=Reg.ReadString('Save', '', ach+'\LTable');
  ForceDirectories(SavePath);
  SolutionsPath:=Reg.ReadString('Save', 'Solutions', SavePath);
  OldSolutionsPath:=SolutionsPath;
  DictPath:=Reg.ReadString('Save', 'Dict', ach+'\Neogramm\Dict.ngr');
  DictFile:=ExtractFileName(DictPath);
  DictPath:=ExtractFilePath(DictPath);
  InterGlossing:=Reg.ReadBool('Save', 'Gloss', True);
  SubstStems:=Reg.ReadBool('Save', 'DictDo', True) and DirectoryExists(DictPath);
  UnconfirmedStems:=Reg.ReadBool('Save', 'DictUnconf', False);
  UploadData:=Reg.ReadBool('Upload', '', False);
  UploadSolv:=Reg.ReadBool('Upload', 'Solutions', False);
  ShowUploadWindow:=Reg.ReadBool('Upload', 'ShowWin', False);
  HostName:=Reg.ReadString('Upload', 'Host', '');
  UserName:=Reg.ReadString('Upload', 'User', '');
  Password:=DecryptPwd(Reg.ReadString('Upload', 'Password', ''), 34960119);
  UploadPath:=Reg.ReadString('Upload', 'Path', '');
  UploadSolvPath:=Reg.ReadString('Upload', 'SolutionsPath', '');
  SetCurrentDir(SavePath);
  sl:=TStringList.Create;
  for i:=0 to Length(Heads)-2 do if ColType[i, ColCounts[i]-1]=ctSolv then SolvTab:=i;
  for i:=0 to Length(Heads)-2 do begin
    TabControl.Tabs.Add(Heads[i]);
    try sl.LoadFromFile(Heads[i]+'.dat') except sl.Clear end;
    Backups[i]:=sl.Text;
    for j:=0 to sl.Count-1 do begin
      p:=MultPos('|', sl[j]);
      if p>1 then begin
        SetLength(Tables[i], Length(Tables[i])+1);
        Tables[i, Length(Tables[i])-1]:=TTable.Create;
        with Tables[i, Length(Tables[i])-1] do begin
          Un:=StrToInt(Copy(sl[j], 1, 2));
          if (Length(Tables[i])>=2) and (Un<Tables[i, Length(Tables[i])-2].Un) then Un:=Tables[i, Length(Tables[i])-2].Un;
          Id:=Copy(sl[j], 3, p-3);
          k:=LastDelimiter('¤', Id);   
          if k>0 then begin
            Nut:=StrToIntDef(Copy(Id, k+1, MaxInt), 255);
            Delete(Id, k, MaxInt);
          end else Nut:=255;
        end;
      end;
      with Tables[i, Length(Tables[i])-1] do begin
        SetLength(HTMLs, Length(HTMLs)+1);
        SetLength(NoNutshell, Length(HTMLs));
        NoNutshell[High(NoNutshell)]:=Copy(sl[j], p+1, 1)='¤';
        if NoNutshell[High(NoNutshell)] then Inc(p);
        for k:=0 to ColCounts[i]-1+Ord(i=SolvTab) do begin
          q:=Pos('|', Copy(sl[j], p+1, MaxInt));
          if q=0 then q:=NoPos;
          HTMLs[High(HTMLs), k]:=StringReplace(StringReplace(Copy(sl[j], p+1, q-1), '\u164?', '¤', [rfReplaceAll]), '\u124?', '|', [rfReplaceAll]);
          if k=ColCounts[i] then HTMLs[High(HTMLs), k]:=StringReplace(HTMLs[High(HTMLs), k], '\u10?', #13#10, [rfReplaceAll]);
          Inc(p, q);
        end;
      end;
    end;
  end;
  for i:=0 to LemBar.ButtonCount-1 do with LemBar.Buttons[i] do if Hint='' then Hint:=' '+Caption+' ';
  SplitIniLine(Reg.ReadString('Environment', 'Appendix', 'Time|Date|Meas|Stars|Texts')+'|', UnitCombo);
  UpdateTextPages;
  SplitIniLine(Reg.ReadString('Environment', 'TextTags', 'h2|p|span|a href="../le.php?$"|a href="?lemma=$"|td|th')+'|', TextTagsCombo);
  try sl.LoadFromFile('Text.txt') except sl.Clear end;
  Backups[High(Heads)]:=sl.Text;
  for i:=0 to Length(TextPages)-1 do begin
    p:=Pos('|'#13#10, sl.Text);
    if p=0 then p:=Length(sl.Text)+1;
    TextPages[i].Text:=StringReplace(Copy(sl.Text, 1, p-1), '\u124?', '|', [rfReplaceAll]);
    sl.Text:=Copy(sl.Text, p+3, MaxInt);
  end;
  sl.Free;
  TabControl.Tabs.Add('Lemizh text');
  TabControlChange(nil);
  FindEdit.Text:=Reg.ReadString('Find', '', '');  {vor Open Tab/ListBox/Row/Col}
  LemFindBtn.Tag:=Min(Max(Reg.ReadInteger('Find', 'Lem', 0), 0), 2);
  LemFindBtnClick(nil);
  CaseCheckBox.Checked:=Reg.ReadBool('Find', 'Case', False);
  AccentCheckBox.Checked:=Reg.ReadBool('Find', 'Accent', True);
  VowelCheckBox.Checked:=Reg.ReadBool('Find', 'Vowel', False);
  OptsFind:=Reg.ReadString('Find', 'Options', '');
  for i:=0 to Length(Heads)-2 do for j:=0 to 3 do ColWidths[i, j]:=Max(Reg.ReadInteger('Columns', IntToStr(i)+IntToStr(j), (StringGrid.Width-20) div ColCounts[i]), 5);
  for i:=0 to Length(Heads)-2 do ListBoxIndex[i]:=Max(Min(Reg.ReadInteger('Open', 'ListBox'+IntToStr(i), Ord(Length(Tables[i])>0)-1), Length(Tables[i])-1), 0);
  TabControl.TabIndex:=Reg.ReadInteger('Open', '', 0);
  TabControlChange(nil);
  try StringGrid.Row:=Reg.ReadInteger('Open', 'Row', 1) except end;
  try StringGrid.Col:=Reg.ReadInteger('Open', 'Col', 0) except end;
  UnitList.ItemIndex:=Min(Max(Reg.ReadInteger('Open', 'Memo', -1)+1, 0), UnitList.Items.Count-1);
  WindowState:=TWindowState(Reg.ReadInteger('Win', 'State', Ord(wsNormal)));
  if WindowState=wsNormal then begin
    Left:=Reg.ReadInteger('Win', 'Left', 70);
    Top:=Reg.ReadInteger('Win', 'Top', 140);
    Width:=Reg.ReadInteger('Win', 'Width', 760);
    Height:=Reg.ReadInteger('Win', 'Height', 420);
  end;
  ListPanel.Width:=Reg.ReadInteger('Win', 'Splitter', 130);
  HTML5Image.Visible:=Reg.ReadBool('Win', 'HTML5', False);
  LemBox.Checked:=Reg.ReadBool('Win', 'LemFont', True);
  NeogrammPath:=Reg.ReadString('Win', 'Neogramm', '');
  SaveCaretBox.Checked:=Reg.ReadBool('Text', 'SaveCaret', True);
  MemoFontSize:=Reg.ReadInteger('Text', 'Fontsize', 0);
  SolvMemo.Font.Size:=Font.Size+Reg.ReadInteger('Solv', 'Fontsize', 0);
  for i:=0 to Length(TextPages)-1 do with TextPages[i] do begin
    if i=0 then st:='' else st:=IntToStr(i-1);
    if SaveCaretBox.Checked then begin
      Caret:=Reg.ReadInteger('Text', 'Caret'+st, Length(Text));
      Selection:=Reg.ReadInteger('Text', 'Sel'+st, 0);
    end else Caret:=Length(Text);
    Header:=Reg.ReadString('Text', 'Header'+st, 'span');
    ColSpan:=Reg.ReadInteger('Text', 'Colspan'+st, 1);
    RowSpan:=Reg.ReadInteger('Text', 'Rowspan'+st, 1);
  end;
  UnitListChange(nil);
  TextTagsComboChange(nil);
  if Reg.ReadBool('Win', 'Multiselect', False) then begin
    MultiSelectBtn.Down:=not MultiSelectBtn.Down;
    MultiSelectBtnClick(nil);
  end;
  for i:=0 to Length(Heads)-2 do begin
    MarkBackups[i]:=Reg.ReadString('Marks', IntToStr(i), '');
    st:=Reg.ReadString('Marks', IntToStr(i)+'c', IntToStr(Length(MarkBackups[i]) div 4))+',';
    q:=1;
    for j:=1 to MarksHi do begin
      p:=Pos(',', st);
      q2:=StrToIntDef(Copy(st, 1, p-1), 0);
      Delete(st, 1, p);
      for k:=q to q+q2-1 do begin
        try p:=HexToInt(Copy(MarkBackups[i], 4*k-3  +j-1, 4)) except p:=MaxInt; end;
        if Length(Tables[i])>p then Tables[i, p].Mark:=j;
      end;
      Inc(q, q2);
      Insert('.', MarkBackups[i], 4*q+j-4);
    end;
  end;
  for i:=0 to High(Symbols) do begin
    SetLength(Symbols[i], Reg.ReadInteger('Symbols', '', 0));
    for j:=0 to Length(Symbols[0])-1 do Symbols[i,j]:=Reg.ReadString('Symbols', IntToStr(j)+SymLetters[i+1], '');
  end;
  CaseMenuBreak:=Reg.ReadBool('Symbols', 'Break', False);
  for i:=0 to High(CaseAbbrs) do begin
    CaseAbbrs[i]:=Reg.ReadString('Cases', IntToStr(i)+'a', '');
    CaseExtds[i]:=Reg.ReadString('Cases', IntToStr(i), '');
  end;
  LemFontCorr:=Reg.ReadInteger('Environment', 'LemCorr', 0);
  Application.HintHidePause:=1000*Reg.ReadInteger('Environment', 'HintHide', 10);
  Reg.Free;
  UpdateCasePopup;
  VLabel.Caption:='Version '+VersionNr(Application.ExeName);
  dt:=CompileTime(Application.ExeName);
  VLabel.Hint:='Compiled '+DateTimeToStr(dt);
  CLabel.Caption:='2008-'+FormatDateTime('yy', dt)+' by';
  for i:=0 to Length(Heads)-2 do CF_Rows[i]:=RegisterClipboardFormat(PChar('Lemizh Table Maker '+Heads[i]));
  Application.OnActivate:=AppActivate;
  TimeLabel.Caption:='Started in '+FloatToStrF((Now-tm)*86400, ffGeneral, 4, 2)+' s';
  MemoTimeLabel.Caption:=TimeLabel.Caption;
end; {FormCreate}

procedure TTableMaker.FormShow(Sender: TObject);
var h: HWnd;
begin
  HTML5Image.Hint:=HTML5Panel.Hint;
  Options.UnconfirmedBox.Checked:=UnconfirmedStems;
  try
    Options.DriveComboBox3.Drive:=DictPath[1];
    Options.DirectoryListBox3.Directory:=DictPath;
    Options.FileListBox3.FileName:=DictFile;
  except end;
  Options.DirectoryListBox1.Update;
  Options.DirectoryListBox2.Update;
  Options.FileListBox3.OnChange:=Options.DirectoryListBox3Change;
  h:=FindWindowEx(0, Handle, 'TTableMaker', nil);
  if (h<>Handle) and (GetWindowTextLength(h)>0) then begin
    SetForegroundWindow(h);
    PostMessage(Handle, WM_Close, 0, 0);
  end;
end; {FormShow}

procedure TTableMaker.AppActivate(Sender: TObject);
begin
  if ShallSave=mrCancel then begin
    Options.UpdateDict(nil);
    Options.PageControlChange(nil);
  end;
end; {AppActivate}

procedure TTableMaker.AppShowHint(var HintStr: string; var CanShow: Boolean; var HintInfo: THintInfo);
var c, r, n: Integer;
    sw: string;
    Files, Marks: array[0..Length(Heads)-1] of string;
begin
  with HintInfo do if HintControl=StringGrid then begin
    StringGrid.MouseToCell(CursorPos.X, CursorPos.Y, c, r);
    if (c>-1) and (r>-1) then begin
      CursorRect:=StringGrid.CellRect(c, r);
      if r>0 then begin
        HintStr:=ThisTable.HTMLs[r-1, c];
        if ColType[TabControl.TabIndex, c]=ctSolv then begin
          if ThisTable.HTMLs[r-1, c+1]='' then HintStr:='Double click to create' else HintStr:='Double click to edit:'#13#10+ThisTable.HTMLs[r-1, c+1];
        end;
      end else if ColType[TabControl.TabIndex, c]=ctSolv then begin
       GetNoOfSolutions(1, Length(ThisTable.HTMLs), UnitCombo.ItemIndex, IdEdit.Text, n, c);
       if n>1 then HintStr:='Double click to sort';
      end;
    end;
  end else if (HintControl=OKBtn) or (HintControl=SaveBtn) then begin
    sw:=SaveWhat(Files, Marks);
    if sw<>'' then HintStr:='Save '+sw+IfThen(HintControl=OKBtn, '; then close') else if HintControl=OKBtn then HintStr:='Close' else HintStr:='Nothing needs to be saved.';
  end else if HintControl=SyntaxLabel then begin
    if SolvPanel.Visible then HintStr:=SolvSyntax else HintStr:=TableSyntax;
  end else if HintControl=MemoSyntaxLabel then HintStr:=MemoSyntax
  else if HintControl=ShortcutsImage then
    if SolvPanel.Visible then HintStr:=MemoShortcuts else HintStr:=AllShortcuts;
end; {AppShowHint}

procedure TTableMaker.OKBtnClick(Sender: TObject);
begin
  ShallSave:=mrYes;
  Close;
end; {OKBtnClick}

procedure TTableMaker.CancelBtnClick(Sender: TObject);
begin
  Close;
end; {CancelBtnClick}

procedure TTableMaker.SaveBtnClick(Sender: TObject);
var i: Integer;
    tm: TDateTime;
    Files, Marks: array[0..Length(Heads)-1] of string;
begin
  tm:=Now;
  for i:=0 to Length(Heads)-1 do MakeFileAndMarks(i, Files[i], Marks[i]);
  Save(True, Files, Marks);
  if ActiveControl=SaveBtn then if Notebook.PageIndex=1 then Memo.SetFocus else if SolvPanel.Visible then SolvMemo.SetFocus else StringGrid.SetFocus;
  TimeLabel.Caption:='Saved in '+FloatToStrF((Now-tm)*86400, ffGeneral, 4, 2)+' s';
  MemoTimeLabel.Caption:=TimeLabel.Caption;
end; {SaveBtnClick}

procedure TTableMaker.MakeFileAndMarks(Nr: Integer; out AFile, AMark: string);
var j, k, l: Integer;
    st: string;
begin
  AFile:='';
  if Nr=Length(Heads)-1 then for j:=0 to UnitList.Items.Count-1 do AFile:=AFile+StringReplace(TextPages[j].Text, '|', '\u124?', [rfReplaceAll])+'|'#13#10
      else for j:=0 to Length(Tables[Nr])-1 do with Tables[Nr, j] do for k:=0 to Length(HTMLs)-1 do begin
    if k=0 then st:=FormatFloat('00', Un)+Id+IfThen(Nut<255, '¤'+IntToStr(Nut)) else st:='';
    for l:=0 to ColCounts[Nr]-1+Ord(Nr=SolvTab) do st:=st+'|'+IfThen((l=0) and NoNutshell[k], '¤')+
      StringReplace(StringReplace(StringReplace(HTMLs[k,l], '¤', '\u164?', [rfReplaceAll]), '|', '\u124?', [rfReplaceAll]), #13#10, '\u10?', [rfReplaceAll]);
    AFile:=AFile+st+#13#10;                                       
  end;
  AMark:='';
  if Nr<Length(Heads)-1 then for k:=1 to MarksHi do begin
    for j:=0 to Length(Tables[Nr])-1 do if Tables[Nr, j].Mark=k then AMark:=AMark+IntToHex(j, 4);
    AMark:=AMark+'.';
  end;
end; {MakeFileAndMarks}

function TTableMaker.SaveWhat(out Files, Marks: array of string): string;
var i: Integer;
    b: Boolean;
begin
  for i:=0 to Length(Heads)-1 do MakeFileAndMarks(i, Files[i], Marks[i]);
  Result:='';
  b:=Copy(Files[Length(Heads)-1], Pos('|', Files[Length(Heads)-1]), MaxInt)=Copy(Backups[Length(Heads)-1], Pos('|', Backups[Length(Heads)-1]), MaxInt);
  for i:=0 to Length(Heads)-1 do if Files[i]<>Backups[i] then Result:=Result+IfThen((i=Length(Heads)-1) and b, Text0)+Heads[i]+', ';
  b:=False;
  for i:=0 to Length(Heads)-2 do if Marks[i]<>MarkBackups[i] then b:=True;
  if b then Result:=Result+'bookmarks, ';
  if OptsChanged then Result:=Result+'options, ';
  Result:=LowerCase(Copy(Result, 1, Length(Result)-2));
end; {SaveWhat}

function TTableMaker.SaveSolvHTMLs(Always: Boolean): string;
  function BackupSolv(Table, SolvId: string): string;
  var p: Integer;
  begin
    p:=Pos('|'+SolvId+'|', Table);
    if p>0 then Result:=Copy(Table, p+Length(SolvId)+2, MaxInt) else Result:='';
    Result:=StringReplace(Copy(Result, 1, Pos(#13#10, Result)-1), '\u10?', #13#10, [rfReplaceAll]);
  end;
var i, j, k, n, p: Integer;
    sl: TStringList;
    st, st2, fname, errmsg, errmsg1: string;
begin
  sl:=TStringList.Create;
  errmsg:='';
  ForceDirectories(SolutionsPath);
  SaveSolv.Gauge.MaxValue:=100*Length(Tables[SolvTab]);
  Result:='';
  for i:=0 to Length(Tables[SolvTab])-1 do with Tables[SolvTab, i] do begin
    p:=Pos(#13#10+FormatFloat('00', Un)+Id+'|', #13#10+Backups[SolvTab]);
    if p>0 then st:=Copy(Backups[SolvTab], p, MaxInt) else st:='';
    n:=48;
    repeat
      p:=Pos(#13#10+Char(n), st);
      if p=0 then Inc(n);
    until (n>57) or (p>0);
    if n<=57 then st:=Copy(st, 1, p+1);
    for j:=0 to Length(HTMLs)-1 do if (HTMLs[j, ColCounts[SolvTab]]<>'') and (Always or (HTMLs[j, ColCounts[SolvTab]]<>BackupSolv(st, HTMLs[j, ColCounts[SolvTab]-1]))) then begin
      SaveSolv.Gauge.Progress:=100*i+100*j div Length(HTMLs);
      if not SaveSolv.Visible then begin
        SaveSolv.Show;
        SaveSolv.Update;
      end;
      sl.Text:=HTMLs[j, ColCounts[SolvTab]];
      errmsg1:='';
      for k:=sl.Count-1 downto 0 do begin
        if InterGlossing then begin
          st2:=TestLemParts(sl[k], False, errmsg1);
          if st2<>'' then sl.Insert(k+1, '<div class="igloss">'+st2+'</div>');
        end;
        sl[k]:=TrimLeft(sl[k]);
      end;
      sl.Text:=HTML4or5(RTFtoUTF8(IncludeAbbrEtc(SubstituteStems(sl.Text, False, Options.Dictionary.Items), errmsg1)));
      for k:=0 to sl.Count-2 do if (Copy(sl[k], Length(sl[k])-5, 6)<>'</div>') and (Copy(sl[k+1], 1, 4)<>'<div') then sl[k]:=sl[k]+'<br>';
      fname:=IntToStr(Un)+'_'+LowerCase(Id)+'_'+HTMLs[j, ColCounts[SolvTab]-1]+'.html';
      sl.SaveToFile(SolutionsPath+'\'+fname);
      StripWhitespace(SolutionsPath+'\'+fname);
      Result:=Result+fname+#1;
      if errmsg1<>'' then errmsg:=errmsg+IntToStr(Un)+'. '+Id+' #'+HTMLs[j, ColCounts[SolvTab]-1]+': '#13#10#13#10+errmsg1;
    end;
  end;
  sl.Free;
  SaveSolv.Hide;
  if errmsg<>'' then ErrorMessageBox('Errors in solutions:'#13#10#13#10+errmsg);
end; {SaveSolvHTMLs}

procedure TTableMaker.Save(All: Boolean; Files, Marks: array of string);
  function ChangeDir(D: string; SL: TStringList): Integer;
  var i: Integer;
  begin
    D:=StringReplace(D, '/', '\', [rfReplaceAll])+IfThen(Copy(D, Length(D), 1)<>'\', '\');
    Result:=0;
    repeat
      i:=Pos('\', D);
      if i>0 then begin
        SL.Add('cd "'+Copy(D, 1, i-1)+'"');
        D:=Copy(D, i+1, 20000);
        Inc(Result);
      end;
    until i=0;
  end;
const Ext: array[False..True] of string = ('dat', 'txt');
var i, j, k: Integer;
    b: Boolean;
    Reg: TRegIniFile;
    st, solv: string;
    sl: TStringList;
    t: TDateTime;
    tosave: array[-1..High(Heads)] of Boolean;
begin
  Reg:=TRegIniFile.Create('SOFTWARE\TESSoft Inc.\LTable');
  Reg.WriteInteger('Open', '', TabControl.TabIndex);
  for i:=0 to Length(Heads)-2 do Reg.WriteInteger('Open', 'ListBox'+IntToStr(i), ListBoxIndex[i]);
  Reg.WriteInteger('Open', 'Row', StringGrid.Row);
  Reg.WriteInteger('Open', 'Col', StringGrid.Col);
  Reg.WriteInteger('Open', 'Memo', UnitList.ItemIndex-1);
  TabControlChanging(nil, b);
  for i:=0 to Length(Heads)-2 do for j:=0 to 3 do Reg.WriteInteger('Columns', IntToStr(i)+IntToStr(j), ColWidths[i, j]);
  Reg.WriteInteger('Win', 'State', Ord(WindowState));
  Reg.WriteInteger('Win', 'Left', Left);
  Reg.WriteInteger('Win', 'Top', Top);
  Reg.WriteInteger('Win', 'Width', Width);
  Reg.WriteInteger('Win', 'Height', Height);
  Reg.WriteInteger('Win', 'Splitter', Splitter.Left);
  Reg.WriteBool('Win', 'Multiselect', MultiSelectBtn.Down);
  Reg.WriteBool('Win', 'HTML5', HTML5Image.Visible);
  Reg.WriteBool('Win', 'LemFont', LemBox.Checked);
  Reg.WriteString('Win', 'Neogramm', NeogrammPath);
  Reg.WriteBool('Text', 'SaveCaret', SaveCaretBox.Checked);
  Reg.WriteInteger('Text', 'Fontsize', MemoFontSize);
  Reg.WriteInteger('Solv', 'Fontsize', SolvMemo.Font.Size-Font.Size);
  for i:=0 to Length(TextPages)-1 do with TextPages[i] do begin
    if i=0 then st:='' else st:=IntToStr(i-1);
    if SaveCaretBox.Checked then begin
      Reg.WriteInteger('Text', 'Caret'+st, Caret);
      Reg.WriteInteger('Text', 'Sel'+st, Selection);
    end;
    Reg.WriteString('Text', 'Header'+st, Header);
    Reg.WriteInteger('Text', 'Colspan'+st, ColSpan);
    Reg.WriteInteger('Text', 'Rowspan'+st, RowSpan);
  end;
  Reg.WriteString('Find', '', FindEdit.Text);
  Reg.WriteInteger('Find', 'Lem', LemFindBtn.Tag);
  Reg.WriteBool('Find', 'Case', CaseCheckBox.Checked);
  Reg.WriteBool('Find', 'Accent', AccentCheckBox.Checked);
  Reg.WriteBool('Find', 'Vowel', VowelCheckBox.Checked);
  Reg.WriteString('Find', 'Options', Options.FindEdit.Text);
  if All then begin
    for i:=0 to High(CaseAbbrs) do begin
      Reg.WriteString('Cases', IntToStr(i)+'a', CaseAbbrs[i]);
      Reg.WriteString('Cases', IntToStr(i), CaseExtds[i]);
    end;
    Reg.EraseSection('Symbols');
    Reg.WriteInteger('Symbols', '', Length(Symbols[0]));
    for j:=0 to Length(Symbols[0])-1 do for i:=0 to High(Symbols) do Reg.WriteString('Symbols', IntToStr(j)+SymLetters[i+1], Symbols[i,j]);
    Reg.WriteBool('Symbols', 'Break', CaseMenuBreak);
    Reg.WriteBool('Save', 'Backup', BackupFiles);
    Reg.WriteString('Save', '', SavePath);
    Reg.WriteString('Save', 'Solutions', SolutionsPath);
    Reg.WriteString('Save', 'Dict', DictPath+DictFile);
    Reg.WriteBool('Save', 'Gloss', InterGlossing);
    Reg.WriteBool('Save', 'DictDo', SubstStems and DirectoryExists(DictPath));
    Reg.WriteBool('Save', 'DictUnconf', UnconfirmedStems);
    Reg.WriteBool('Upload', '', UploadData);
    Reg.WriteBool('Upload', 'Solutions', UploadSolv);
    Reg.WriteBool('Upload', 'ShowWin', ShowUploadWindow);
    Reg.WriteString('Upload', 'Host', HostName);
    Reg.WriteString('Upload', 'User', UserName);
    Reg.WriteString('Upload', 'Password', EncryptPwd(Password, 34960119));
    Reg.WriteString('Upload', 'Path', UploadPath);
    Reg.WriteString('Upload', 'SolutionsPath', UploadSolvPath);
    for i:=0 to Length(Heads)-2 do begin
      MarkBackups[i]:=Marks[i];
      Reg.WriteString('Marks', IntToStr(i), StringReplace(Marks[i], '.', '', [rfReplaceAll]));
      st:='';
      for j:=1 to MarksHi do begin
        k:=Pos('.', Marks[i]);
        Delete(Marks[i], 1, k);
        st:=st+','+IntToStr(k div 4);
      end;
      Reg.WriteString('Marks', IntToStr(i)+'c', Copy(st, 2, MaxInt));
    end;
    OptsChanged:=False;
    SetCurrentDir(SavePath);
    solv:=SaveSolvHTMLs(SolutionsPath<>OldSolutionsPath);
    OldSolutionsPath:=SolutionsPath;
    sl:=TStringList.Create;
    tosave[-1]:=False;
    for i:=0 to Length(Heads)-1 do if Files[i]<>Backups[i] then begin
      if BackupFiles then CopyFile(PChar(Heads[i]+'.'+Ext[i=Length(Heads)-1]), PChar(Heads[i]+'.bak'), False);
      sl.Text:=Files[i];
      sl.SaveToFile(Heads[i]+'.'+Ext[i=Length(Heads)-1]);
      Backups[i]:=Files[i];
      tosave[i]:=True;
      tosave[-1]:=True;
    end else tosave[i]:=False;
    sl.Free;
    if (UploadData and tosave[-1]) or (UploadSolv and (solv<>'')) then begin        
      sl:=TStringList.Create;
      sl.Add('ftp -s:upload.bat');
      sl.Add('goto done');
      sl.Add('open "'+HostName+'"');
      sl.Add(UserName);
      sl.Add(Password);
      sl.Add('ascii');
      if UploadData and tosave[-1] then begin
    //    SetCurrentDir(SavePath);
        k:=ChangeDir(UploadPath, sl);
        for i:=0 to Length(Heads)-1 do if tosave[i] then sl.Add('put "'+Heads[i]+'.'+Ext[i=Length(Heads)-1]+'"');
        if UploadSolv and (solv<>'') then for i:=1 to k do sl.Add('cd ..');
      end;
      if UploadSolv and (solv<>'') then begin
        SetCurrentDir(SolutionsPath);
        ChangeDir(UploadSolvPath, sl);
        repeat
          i:=Pos(#1, solv);
          if i>0 then begin
            sl.Add('put "'+Copy(solv, 1, i-1)+'"');
            solv:=Copy(solv, i+1, 20000);
          end;
        until i=0;
      end;
      sl.Add('bye');
      sl.Add(':done');
      sl.Add('del upload.bat');
      t:=Now;
      repeat i:=FileCreate('upload.bat') until (i>-1) or (Now>t+0.001);
      FileClose(i);
      try sl.SaveToFile('upload.bat') except end;
      sl.Free;
      WinExec('upload.bat', Ord(ShowUploadWindow));
    end;
    Reg.WriteInteger('Environment', 'LemCorr', LemFontCorr);
    Reg.WriteInteger('Environment', 'HintHide', Application.HintHidePause div 1000);
    st:='';
    for i:=18 to UnitCombo.Items.Count-1 do st:=st+StringReplace(UnitCombo.Items[i], '|', '\u124?', [rfReplaceAll])+'|';
    Reg.WriteString('Environment', 'Appendix', Copy(st, 1, Length(st)-1));
    st:='';
    for i:=0 to TextTagsCombo.Items.Count-1 do st:=st+StringReplace(TextTagsCombo.Items[i], '|', '\u124?', [rfReplaceAll])+'|';
    Reg.WriteString('Environment', 'TextTags', Copy(st, 1, Length(st)-1));
  end;
  Reg.Free;
end; {Save}

procedure TTableMaker.FormClose(Sender: TObject; var Action: TCloseAction);
var i, j: Integer;
    Files, Marks: array[0..Length(Heads)-1] of string;
    st: string;
begin
  st:=SaveWhat(Files, Marks);
  if (ShallSave=mrCancel) and (st='') then ShallSave:=mrNo;
  if (ShallSave=mrCancel) and (st=Text0+LowerCase(Heads[Length(Heads)-1])) then ShallSave:=mrYes;
  if ShallSave=mrCancel then ShallSave:=TangoMessageBox('Save changes ('+st+')?', mtConfirmation, [mbYes, mbNo, mbCancel], '');
  if ShallSave=mrCancel then Action:=caNone else begin
    Save(ShallSave=mrYes, Files, Marks);
    for i:=0 to Length(Tables)-1 do for j:=0 to Length(Tables[i])-1 do Tables[i, j].Free;
  end;
end; {FormClose}

{------------------------------------------------ Window -----------------------------------------------------------}

procedure TTableMaker.SplitterMoved(Sender: TObject);
begin
  ListBox.Repaint;
end; {SplitterMoved}

procedure TTableMaker.ShowOpts(TabNr: Integer);
var i, j: Integer;
begin
  Options.LemBox.Checked:=LemBox.Checked;
  with Options do begin
    PageControl.ActivePage:=PageControl.Pages[TabNr];
    PageControlChange(nil);
    with CaseGrid do for i:=0 to High(CaseAbbrs) do begin
      Cells[1, i+1]:=CaseAbbrs[i];
      Cells[2, i+1]:=CaseExtds[i];
    end;
    SymbGrid.RowCount:=Max(Length(Symbols[0]), 1)+1;
    with SymbGrid do for i:=0 to High(Symbols) do for j:=0 to Length(Symbols[0])-1 do Cells[i, j+1]:=Symbols[i, j];
    MenuBreakBox.Checked:=CaseMenuBreak;
    UpdateButtons;
    BackupFilesBox.Checked:=BackupFiles;
    try
      DriveComboBox1.Drive:=SavePath[1];
      DirectoryListBox1.Directory:=SavePath;
    except end;
    try
      DriveComboBox2.Drive:=SolutionsPath[1];
      DirectoryListBox2.Directory:=SolutionsPath;
    except end;
    try
      DriveComboBox3.Drive:=DictPath[1];
      DirectoryListBox3.Directory:=DictPath;
      FileListBox3.FileName:=DictPath+DictFile;
    except FileListBox3.ItemIndex:=-1 end;
    UploadCheckBox.Checked:=UploadData;
    UploadSolvCheckBox.Checked:=UploadSolv;
    ShowUploadCheckBox.Checked:=ShowUploadWindow;
    HostEdit.Text:=HostName;
    UserEdit.Text:=UserName;
    PasswordEdit.Text:=Password;
    UploadPathEdit.Text:=UploadPath;
    UploadSolvPathEdit.Text:=UploadSolvPath;
    GlossingBox.Checked:=InterGlossing;
    SubstStemsBox.Checked:=SubstStems and DirectoryExists(DictPath);
    GridSetEditText(nil, 0, 0, '');
    SolvSizeEdit.Value:=SolvMemo.Font.Size;
    MemoSizeEdit.Value:=MemoFontSize+TableMaker.Font.Size+1;
    LemFontCorrEdit.Value:=LemFontCorr;
    HintHideEdit.Value:=Application.HintHidePause div 1000;
    Appendices.Lines.Clear;
    for i:=18 to UnitCombo.Items.Count-1 do Appendices.Lines.Add(UnitCombo.Items[i]);
    TextTags.Lines.Assign(TextTagsCombo.Items);
    if OptsFind<>'' then Options.FindEdit.Text:=OptsFind;
    OptsFind:='';
    if ShowModal=idOK then begin
      OptsChanged:=True;
      for i:=0 to High(CaseAbbrs) do with CaseGrid do begin
        CaseAbbrs[i]:=LowerCase(Cells[1, i+1]);
        CaseExtds[i]:=Cells[2, i+1];
      end;
      for i:=0 to 2 do SetLength(Symbols[i], SymbGrid.RowCount-1);
      with SymbGrid do for j:=0 to RowCount-2 do for i:=0 to High(Symbols) do
        Symbols[i, j]:=Cells[i, j+1];
      CaseMenuBreak:=MenuBreakBox.Checked;
      UpdateCasePopup;
      BackupFiles:=BackupFilesBox.Checked;
      SavePath:=DirectoryListBox1.Directory;
      SolutionsPath:=DirectoryListBox2.Directory;
      DictPath:=DirectoryListBox3.Directory+'\';
      DictFile:=ExtractFileName(FileListBox3.FileName);
      UploadData:=UploadCheckBox.Checked;
      UploadSolv:=UploadSolvCheckBox.Checked;
      ShowUploadWindow:=ShowUploadCheckBox.Checked;
      HostName:=HostEdit.Text;
      UserName:=UserEdit.Text;
      Password:=PasswordEdit.Text;
      UploadPath:=UploadPathEdit.Text;
      UploadSolvPath:=UploadSolvPathEdit.Text;
      InterGlossing:=GlossingBox.Checked;
      SubstStems:=SubstStemsBox.Checked and DirectoryExists(DictPath);
      UnconfirmedStems:=UnconfirmedBox.Checked;
      SolvMemo.Font.Size:=SolvSizeEdit.Value;
      MemoFontSize:=MemoSizeEdit.Value-TableMaker.Font.Size-1;
      SetLemFont(Memo.Font, True, MemoFontSize);
      LemFontCorr:=LemFontCorrEdit.Value;
      Application.HintHidePause:=1000*HintHideEdit.Value;
      for i:=UnitCombo.Items.Count-1 downto 18 do UnitCombo.Items.Delete(i);
      for i:=0 to Min(Appendices.Lines.Count, MaxAppendices)-1 do if Trim(Appendices.Lines[i])<>'' then UnitCombo.Items.Add(Appendices.Lines[i]);
      for i:=0 to Length(Heads)-2 do for j:=0 to Length(Tables[i])-1 do with Tables[i,j] do Un:=Min(Un, UnitCombo.Items.Count-1);
      if TabControl.TabIndex<High(Heads) then ListBoxClick(nil);
      ListBox.Invalidate;
      StringGrid.Invalidate;
      UpdateTextPages;
      TextTagsCombo.Items.Assign(TextTags.Lines);
    end else FormShow(nil);
  end;
end; {ShowOpts}

procedure TTableMaker.OptsBtnClick(Sender: TObject);
begin
  ShowOpts(TButton(Sender).Tag);
end; {OptsBtnClick}

procedure TTableMaker.LinkImageClick(Sender: TObject);
begin
  if Sender=LemizhImage then OpenHTTP('https://lemizh.conlang.org/') else OpenHTTP('https://creativecommons.org/licenses/by-nc-sa/4.0/');
end; {LinkImageClick}

procedure TTableMaker.TabControlChanging(Sender: TObject; var AllowChange: Boolean);
var i: Integer;
begin
  StringGrid.EditorMode:=False;
  if Notebook.PageIndex=0 then for i:=0 to StringGrid.ColCount-1 do ColWidths[TabControl.TabIndex, i]:=StringGrid.ColWidths[i];
end; {TabControlChanging}

procedure TTableMaker.TabControlChange(Sender: TObject);
var i: Integer;
begin
  SolvOKBtnClick(nil);
  Notebook.PageIndex:=Ord(TabControl.TabIndex>High(Heads)-1);
  if Notebook.PageIndex=0 then begin
    StringGrid.ColCount:=ColCounts[TabControl.TabIndex];
    for i:=0 to StringGrid.ColCount-1 do StringGrid.ColWidths[i]:=ColWidths[TabControl.TabIndex, i];
    for i:=0 to StringGrid.ColCount-1 do StringGrid.Cells[i, 0]:=Capt[ColType[TabControl.TabIndex, i]];
    ListBox.Clear;
    for i:=0 to Length(Tables[TabControl.TabIndex])-1 do ListBox.Items.Add('');
    ListBoxModified(ListBoxIndex[TabControl.TabIndex]);
  end else if Sender<>FindEdit then ActiveControl:=Memo;
  EnableLemBar(StringGrid.Col);
  SetCaption;
end; {TabControlChange}

procedure TTableMaker.ListBoxModified(Focus: Integer);
begin
  GridPanel.Visible:=ListBox.Items.Count>0;
  DelTableBtn.Enabled:=GridPanel.Visible;
  if GridPanel.Visible then ListBox.ItemIndex:=Focus else EnableLemBar(0);
  NrTables.Caption:='('+IntToStr(ListBox.Items.Count)+' table'+Copy('s', 1, Ord(ListBox.Items.Count<>1))+')';
  ListBoxClick(nil);
  ListBox.Invalidate;
end; {ListBoxModified}

procedure TTableMaker.EnableLemBar(Col: Integer);
var b: Boolean;
    i: Integer;
begin
  for i:=0 to LemBar.ButtonCount-1 do with LemBar.Buttons[i] do if Copy(Name, 1, 4)='Tool' then begin
    b:=(((ActiveControl=StringGrid) and (ColType[TabControl.TabIndex, Col]=ctLem)) or (ActiveControl=Memo)
      or ((ActiveControl=FindEdit) and (LemFindBtn.Tag=1))) xor (Name='ToolLem');
    if Copy(Name, 5, 6)='Button' then Indeterminate:=not b else TButton(LemBar.Buttons[i]).Font.Color:=clActive[b];
    Enabled:=(ActiveControl=StringGrid) or (ActiveControl=Memo) or (ActiveControl=FindEdit) or (ActiveControl=SolvMemo);
  end;
  ToolNut.Enabled:=ActiveControl=StringGrid;
  ToolNonut.Enabled:=ToolNut.Enabled;
end; {EnableLemBar}

{------------------------------------------------ ListBox & Co -----------------------------------------------------------}

procedure TTableMaker.ListBoxClick(Sender: TObject);
var i, j: Integer;
    b: Boolean;
begin
  SolvOKBtnClick(nil);
  if (ListBox.Items.Count>0) and (ListBox.ItemIndex>-1) then with ThisTable do begin
    ListBoxIndex[TabControl.TabIndex]:=ListBox.ItemIndex;
    UnitCombo.ItemIndex:=Un;
    IdEdit.Text:=Id;
    NutCombo.ItemIndex:=ShortInt(Nut)+1;
    NutComboChange(nil);
    b:=True;
    for i:=0 to Length(Tables[TabControl.TabIndex])-1 do with Tables[TabControl.TabIndex, i] do
      if (i<>ListBox.ItemIndex) and (Un=UnitCombo.ItemIndex) and (Trim(LowerCase(Id))=Trim(LowerCase(IdEdit.Text))) then b:=False;
    IdEdit.Color:=clSuccess[b, TabControl.TabIndex<>SolvTab];
    StringGrid.RowCount:=Max(Length(HTMLs), 1)+1;
    for i:=0 to Length(HTMLs)-1 do for j:=0 to 3 do StringGrid.Cells[j, i+1]:=HTMLs[i, j];
    StringGrid.Row:=StringGrid.RowCount-1;
    StringGrid.Col:=0;
  end;
  UpTableBtn.Enabled:=DelTableBtn.Enabled and (ListBox.ItemIndex>0) and (Tables[TabControl.TabIndex, ListBox.ItemIndex-1].Un>=ThisTable.Un);
  DownTableBtn.Enabled:=DelTableBtn.Enabled and (ListBox.ItemIndex<ListBox.Items.Count-1) and (ListBox.ItemIndex>-1) and (Tables[TabControl.TabIndex, ListBox.ItemIndex+1].Un<=ThisTable.Un);
  UpdateTableButtons;
  SetCaption;
end; {ListBoxClick}

procedure TTableMaker.ListBoxDblClick(Sender: TObject);
begin
  with ThisTable do if Sender=ListBox then Mark:=Ord(Mark=0) else Mark:=TMenuItem(Sender).Tag*Ord(Mark<>TMenuItem(Sender).Tag);
  ListBox.Invalidate;
end; {ListBoxDblClick}

procedure TTableMaker.ListBoxDrawItem(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
var i, t: Integer;
    b: Boolean;
begin
  t:=TabControl.TabIndex;
  if (t<Length(Heads)-1) and (Index<Length(Tables[t])) then with ListBox.Canvas, Tables[t, Index] do begin
    b:=True;
    for i:=0 to Length(Tables[t])-1 do
      if (i<>Index) and (Tables[t, i].Un=Un) and (Trim(LowerCase(Tables[t, i].Id))=Trim(LowerCase(Id))) then b:=False;
    if odSelected in State then Font.Color:=clSuccess[b, t<>SolvTab] else Brush.Color:=clSuccess[b, t<>SolvTab];
    FillRect(Rect);
    TextOut(Rect.Left+38, Rect.Top, Id);
    if (Index=0) or (Tables[t, Index-1].Un<>Un) then TextOut(Rect.Left+30-TextWidth(UnitCombo.Items[Un]), Rect.Top, UnitCombo.Items[Un]);
    if Mark>0 then ImageList.Draw(ListBox.Canvas, Rect.Right-16, (Rect.Top+Rect.Bottom) div 2 -8, 20+Mark);
  end;
end; {ListBoxDrawItem}

procedure TTableMaker.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if ActiveControl<>StringGrid then
    if ((Key=70{F}) and (Shift=[ssCtrl])) or ((Key=VK_F3) and (ActiveControl<>FindEdit)) or ((Key=VK_F4) and (Notebook.PageIndex=0)) or (Key in [VK_F9, VK_F10]) then begin
      StringGridKeyDown(Sender, Key, Shift);
      Key:=0;
    end;          
end; {FormKeyDown}

procedure TTableMaker.NewTableBtnClick(Sender: TObject);
var i, p: Integer;
    aux: TTable;
begin
  SolvOKBtnClick(nil);
  if ListBox.Items.Count=0 then p:=0 else p:=ListBox.ItemIndex+1;
  SetLength(Tables[TabControl.TabIndex], Length(Tables[TabControl.TabIndex])+1);
  ListBox.Items.Add('');
  aux:=TTable.Create;
  with aux do begin
    SetLength(HTMLs, 1);
    SetLength(NoNutshell, 1);
    NoNutshell[0]:=False;
    if p=0 then Un:=0 else Un:=Tables[TabControl.TabIndex, p-1].Un;
    Nut:=255;
    Mark:=0;
  end;
  for i:=ListBox.Items.Count-1 downto p+1 do Tables[TabControl.TabIndex, i]:=Tables[TabControl.TabIndex, i-1];
  Tables[TabControl.TabIndex, p]:=aux;
  ListBoxModified(p);
  ActiveControl:=UnitCombo;
end; {NewTableBtnClick}

procedure TTableMaker.DelTableBtnClick(Sender: TObject);
var i, l: Integer;
begin                                          
  if TangoMessageBox('Delete the table ‘'+UnitCombo.Text+IfThen(IdEdit.Text<>'', '. '+IdEdit.Text)+'’?', mtConfirmation, [mbYes, mbNo], '')=idYes then begin
    SolvOKBtnClick(nil);
    l:=ListBox.ItemIndex;
    if l=ListBox.Items.Count-1 then Dec(l);
    ThisTable.Free;
    for i:=ListBox.ItemIndex to ListBox.Items.Count-2 do Tables[TabControl.TabIndex, i]:=Tables[TabControl.TabIndex, i+1];
    SetLength(Tables[TabControl.TabIndex], Length(Tables[TabControl.TabIndex])-1);
    ListBox.Items.Delete(0);
    ListBoxModified(l);
  end;
end; {DelTableBtnClick}

procedure TTableMaker.ExchangeTables(I, J: Integer);
var aux: TTable;
begin
  aux:=Tables[TabControl.TabIndex, I];
  Tables[TabControl.TabIndex, I]:=Tables[TabControl.TabIndex, J];
  Tables[TabControl.TabIndex, J]:=aux;
end; {ExchangeTables}

procedure TTableMaker.UpDownTableBtnClick(Sender: TObject);
begin
  SolvOKBtnClick(nil);
  ExchangeTables(ListBox.ItemIndex, ListBox.ItemIndex+TComponent(Sender).Tag);
  ListBoxModified(ListBox.ItemIndex+TComponent(Sender).Tag);
end; {UpDownTableBtnClick}

{------------------------------------------------ Grid & Co -----------------------------------------------------------}

procedure TTableMaker.UnitOrIdChange(Sender: TObject);
var s, p: Integer;
begin
  SolvOKBtnClick(nil);
  ThisTable.Id:=IdEdit.Text;
  s:=Sign(UnitCombo.ItemIndex-ThisTable.Un);
  p:=ListBox.ItemIndex;
  if s<>0 then begin
    while (p+s>=0) and (p+s<=ListBox.Items.Count-1) and (s*(UnitCombo.ItemIndex-Tables[TabControl.TabIndex, p+s].Un)>0) do begin
      ExchangeTables(p, p+s);
      Inc(p, s);
    end;
    Tables[TabControl.TabIndex, p].Un:=UnitCombo.ItemIndex;
  end;
  ListBoxModified(p);
end; {UnitOrIdChange}

procedure TTableMaker.NutComboChange(Sender: TObject);
begin
  ThisTable.Nut:=NutCombo.ItemIndex-1;
  NutCombo.Color:=NutColors[NutCombo.ItemIndex>0, True];
  CopyNutshellBtn.Enabled:=NutCombo.ItemIndex>0;
  StringGrid.Invalidate;
end; {NutComboChange}

procedure TTableMaker.HTML5ImageDblClick(Sender: TObject);
begin
  HTML5Image.Visible:=not HTML5Image.Visible;
end; {HTML5ImageDblClick}

const ActAbbr = '<abbr class="gloss" title="agent">a</abbr>';
      old: array[0..11] of string = ('&', '<', '>', 'a]', '(', ')', '[', ']', '{', '}', #13#10, #9);
      new: array[0..11] of string = ('&amp;', '&lt;', '&gt;', ActAbbr+'</sup>', '<span lang="en-GB">', '</span>', '<sup>', '</sup>', '<mark>', '</mark>', '<br>', ' ');

function TTableMaker.IncludeSymbols(S: string): string;
var i, j, k: Integer;
    code: array[0..1] of string;
begin
  for i:=0 to Length(Symbols[0])-1 do if Symbols[tsyCode, i]<>'' then begin
    code[0]:=Symbols[tsyCode, i];  code[1]:=code[0];
    for j:=0 to 2 do code[1]:=StringReplace(code[1], old[j], new[j], [rfReplaceAll]);
    k:=Ord(code[0]<>code[1]);
    for j:=0 to k do if Symbols[tsyHint, i]='' then S:=StringReplace(S, code[j], Symbols[tsyOut, i], [rfReplaceAll])
      else S:=StringReplace(S, code[j], '<abbr title="'+Symbols[tsyHint, i]+'">'+Symbols[tsyOut, i]+'</abbr>', [rfReplaceAll]);
  end;
  Result:=S;
end; {IncludeSymbols}

function TTableMaker.IncludeBracketsEtc(S: string; Symbols: Boolean): string;
var i, p: Integer;
begin
  repeat
    p:=Max(Pos('{`}', S), Pos('{´}', S));
    if p>0 then begin
      Delete(S, p, 3);
      repeat Dec(p) until (p=0) or (Pos(S[p], 'àèÌìòÒùÙáéýíóÓúÚ ')>0);
      if (p>0) and (S[p]<>' ') then S:=Copy(S, 1, p-1)+'{'+S[p]+'}'+Copy(S, p+1, MaxInt);
    end;
  until p<=0;
  for i:=0 to High(old) do begin
    if i in [4..9] then S:=StringReplace(S, '\'+old[i], #31+IntToStr(i), [rfReplaceAll]);
    S:=StringReplace(S, old[i], new[i], [rfReplaceAll]);
    if i in [4..9] then S:=StringReplace(S, #31+IntToStr(i), old[i], [rfReplaceAll]);
  end;
  if Symbols then Result:=IncludeSymbols(S) else Result:=S;
end; {IncludeBracketsEtc}

function TTableMaker.IncludeAbbrEtc(S: string; var ErrMsg: string): string;
  procedure StringReplaceEx(OldPattern, NewPattern: string);
  var i, l: Integer;
  begin
    l:=Length(OldPattern);
    for i:=Length(S)-l+1 downto 1 do if (Copy(S, i, l)=OldPattern) and ((i=Length(S)-l+1) or not (S[i+l] in ['A'..'Z', 'a'..'z', 'À'..'Ö', 'Ø'..'ö', 'ø'..'ÿ']))
      then S:=Copy(S, 1, i-1)+'-<abbr class="gloss" title="'+NewPattern+'</abbr>'+Copy(S, i+l, MaxInt);
  end;
  procedure CaseReplace(Abbr, Extd: string; Sec: Boolean);
  const bo: array[0..4] of string = ('', '(', '[', '{', '<mark>');
        bc: array[0..4] of string = ('', ')', ']', '}', '</mark>');
  var i, j: Integer;
  begin
    if Sec then for i:=32 to 33 do for j:=0 to 4 do StringReplaceEx('-'+bo[j]+CaseAbbrs[i]+bc[j]+Abbr, bo[j mod 4]+CaseExtds[i]+bc[j mod 4]+' '+Extd+'">'+bo[j]+CaseAbbrs[i]+bc[j]+Abbr);
    StringReplaceEx('-'+Abbr, Extd+'">'+Abbr);
  end;
var i, j: Integer;
begin
  for i:=0 to High(CaseAbbrs) do if CaseAbbrs[i]<>'' then begin
    if Pos(CaseAbbrs[i]+'/', S)>0 then for j:=0 to High(CaseAbbrs) do CaseReplace(CaseAbbrs[i]+'/'+CaseAbbrs[j], CaseExtds[i]+' or '+CaseExtds[j], True);
    CaseReplace(CaseAbbrs[i], CaseExtds[i], i<32);
  end;
  S:=StringReplace(S, '-\', '-', [rfReplaceAll]);
  repeat
    i:=Pos('<lem>', S);
    j:=Pos('</lem>', S);
    if j=0 then j:=Length(S)+1;
    if i>0 then S:=Copy(S, 1, i-1)+'<span lang="x-lm"'+LemTranscript(Copy(S, i+5, j-i-5), Symbols, ErrMsg)+'>'
      +IncludeBracketsEtc(Copy(S, i+5, j-i-5), False)+'</span>'+Copy(S, j+6, MaxInt);
  until i=0;
  Result:=IncludeSymbols(S);
end; {IncludeAbbrEtc}

function TTableMaker.TestLemParts(S: string; Lemizh: Boolean; var ErrMsg: string): string;
var p, q: Integer;
begin
  Result:='';
  if Lemizh then Result:=Glossing(S, ErrMsg) else repeat
    p:=Pos('<lem>', S);
    if p>0 then begin
      q:=Pos('</lem>', S);
      if q=0 then begin
        ErrMsg:=ErrMsg+'Syntax error in "'+Copy(S, p+5, MaxInt)+'": <lem> tag not closed.'#13#10#13#10;
        q:=Length(S);
      end else If MultPos(' .,:;–—', Copy(S, p+5, q-p-5))<NoPos then
        if p=1 then Result:=Glossing(Copy(S, p+5, q-p-5), ErrMsg);
      S:=Copy(S, q+6, MaxInt);
    end;
  until p=0;
end; {TestLemParts}

function TTableMaker.Glossing(S: string; var ErrMsg: string): string;
const punctArray = '~^<>‹›«»–—{}'#1#2#3#4#5#6;          {cave ' '+punctArrayNew[2*i] -> punctArrayNew[2*i] gegen Ende der function Glossing; außerdem pctend:=..., ellipsis for numbers}
      punctArrayNew: array[1..18] of string = ('\u8211?','\u8212?','\u8216?','\u8217?','\u8220?','\u8221?','\u8216?\u8220?','\u8217?\u8221?','\u8229?','\u8230?', '<mark>','</mark>','(',')','[',']','{','}');
      vowels = 'aeyioOuUàèÌìòÒùÙáéýíóÓúÚ';
var err: Integer;
    levels: array of Integer;
    mainpred: array of Boolean;
    act, emact: Boolean;

  function Pronoun(pt, pl: Integer): string;
    function WordsBack(bck: Integer): string;
    begin
      Result:=IntToStr(bck)+' word'+Copy('s', 1, Ord(bck>1))+' back';
    end;
  const previous=' of previous sentence';
  var pts, pls, ref: string;
      back, i, pltarget: Integer;
      pastmainpred: Boolean;     
  begin
    pltarget:=levels[High(levels)]-pl;
    back:=High(levels)+1;
    pastmainpred:=False;
    for i:=1 to pt do repeat
      Dec(back);
      if (back<High(mainpred)) and (mainpred[back+1]) then pastmainpred:=True;
      if levels[back]<pltarget then err:=-1;
    until (back<0) or (levels[back]=pltarget);
    case pltarget of
      0: ref:='to parole'+IfThen(pt=2, previous);
      1: if (back=-1) or mainpred[back] then ref:='to main predicate'+IfThen(pastmainpred, previous) else ref:=WordsBack(High(levels)-back);
      else ref:=WordsBack(High(levels)-back);
    end;
    pts:=StringOfChar('I', pt);
    if pl=0 then pls:='' else pls:='\u8722?'+IntToStr(pl);
    Result:='<abbr class="gloss" title="pronoun type '+pts+' level n'+pls+IfThen(S[1]<>'—', ' (refers '+ref+')')+'">P'+pts+'<sub>n'+pls+'</sub></abbr>';
    if (pltarget<0) and (S[1]<>'—') then err:=9 else if (pt=2) and (pltarget>1) and (back<0) then err:=10;
  end; {Pronoun}

  function GlossWord(wrd: string): string;
    procedure BracketsInEnding(APos: Integer; ASearch: string; AEm: Boolean; var Aq, Aqb: Integer);
    begin
      if APos<Length(wrd) then Aq:=Pos(wrd[APos+1], ASearch+IfThen(AEm, '{')+#1#3#5);
      if Aq>Length(ASearch) then
        if (APos+2<Length(wrd)) and (Pos(wrd[APos+3], IfThen(AEm, '}')+#2#4#6)=Aq-Length(ASearch)) then begin
          Aqb:=Aq-Length(ASearch);
          Aq:=Pos(wrd[APos+2], ASearch);
          if Aq=0 then Aqb:=0;
        end else Aq:=0;
    end; {BracketsInEnding}
    function Recurring(numer, denom: Comp; var err: Integer): string;
    var i, j, k, p: Integer;
        c, d: Comp;
    begin
      d:=denom/Hcd(numer, denom);
      i:=-1;
      repeat
        Inc(i);
        c:=Hcd(d, 10);
        d:=d/c;
      until (i>8) or (c=1);
      j:=0;
      if (d>1) and (c=1) then repeat Inc(j) until (j>8) or ((Round(IntPower(10, j)-1) mod Round(d))=0);
     Result:=FormatFloat(',#################0.##################', numer/denom);
     if (c=1) and (i<=8) and (j<=8) then begin
       p:=Pos('.', Result);
       Result:=Copy(Result, 1, p+i+j);
       for k:=p+i+j+1 downto p+i+2 do Insert('\u773?', Result, k);
     end else err:=15;
    end; {Recurring}
  const pauses: array[False..True, False..True] of string = (('', '(!)'), ('.', '!'));
  var i, j, k, l, ncases, pct, pctatend, p, q, qb, r, rb, pp, comma, recur: Integer;
      a: array[1..2] of Byte;
      b, firstrep, isword, femact: Boolean;
      xwrd, ps: string;
      xps: array[-1..1] of string;
      syms: Array of Byte;
      numer, rdiff, up: Extended;
  begin
    Result:='';
    i:=1;
    while (i<=Length(wrd)) and (wrd[i]=' ') do Inc(i);
    wrd:=Copy(wrd, i, MaxInt);
    for j:=1 to Length(wrd) do for i:=0 to Length(Symbols[0])-1 do if Pos(Symbols[tsyCode, i], wrd)=j then begin
      wrd:=Copy(wrd, 1, j-1)+#29+Copy(wrd, j+Length(Symbols[tsyCode, i]), MaxInt);
      SetLength(syms, Length(syms)+1);
      syms[High(syms)]:=i;
    end;
    wrd:=StringReplace(StringReplace(wrd, #28#29, '', []), #28, '', []);
    if Copy(wrd, 1, 1)='—' then Inc(levels[High(levels)]);
    if wrd='' then pct:=0 else begin
      if wrd[Length(wrd)]=' ' then pct:=0 else if wrd[Length(wrd)] in [',', ';'] then pct:=1 else pct:=2;
      if wrd[Length(wrd)] in [' ', ',', '.', ';', ':'] then Result:=pauses[pct=2, (wrd[Length(wrd)] in [';', ':']) or (Copy(wrd, Length(wrd)-1, 2)='  ')]+' '
         else pct:=-1;
    end;
    if pct>-1 then wrd:=Copy(wrd, 1, Length(wrd)-1);
    if err=0 then begin
      i:=Length(wrd);
      while (i>0) and (wrd[i]=' ') do Dec(i);
      wrd:=Copy(wrd, 1, i);
    end;
    p:=Max(Pos('{`}', wrd), Pos('{´}', wrd));
    femact:=MultPos('àèÌìòÒùÙáéýíóÓúÚ', Copy(wrd, 1, p))<NoPos;
    if p>0 then Delete(wrd, p, 3);
    b:=False;  xwrd:=wrd;
    for i:=1 to Length(wrd) do case wrd[i] of
      '(': b:=True;
      ')': b:=False;
      else if b then xwrd[i]:=#31;
    end;
    if wrd<>'' then pctatend:=Pos(wrd[Length(wrd)], punctArray) else pctatend:=0;
    if Odd(pctatend) or (pctatend in [2, 10]) then pctatend:=0;
    ncases:=Ord(levels[High(levels)]=1); {number of handled case endings, including outer case (even for main predicate)}
    if ncases=1 then a[1]:=0;
    firstrep:=True;  isword:=True;
    repeat
      i:=LastDelimiter(vowels, xwrd);
      if firstrep then isword:=i>0;
      Inc(ncases);  p:=0;  q:=0;  qb:=0;  r:=0;  rb:=0;
      if i>0 then begin
        p:=Pos(wrd[i], vowels)-1;
        BracketsInEnding(i, 'lRr', False, q, qb);
        BracketsInEnding(i+Ord(q>0)+2*Ord(qb>0), 'nm', True, r, rb);
      end;
      ps:=Copy(wrd, i+1+Ord(q>0)+2*Ord(qb>0)+Ord(r>0)+2*Ord(rb>0), MaxInt);
      if ps='|' then Dec(ncases);
      for k:=1 downto 0 do begin
        if k=0 then j:=1 else j:=Length(ps);
        while (j>=1) and (j<=Length(ps)) and (Pos(ps[j], punctArray+'-'#29)>0) do Inc(j, 1-2*k);
        xps[k]:=Copy(ps, 1+j*k{1, j+1}, j*(1-k)-1+k*NoPos{j-1, high});
        for l:=1 to Length(punctArray) do xps[k]:=StringReplace(xps[k], punctArray[l], punctArrayNew[l], [rfReplaceAll]);
        l:=Length(syms)-1;
        repeat
          pp:=LastDelimiter(#29, xps[k]);
          if pp>0 then xps[k]:=Copy(xps[k], 1, pp-1)+Symbols[tsyCode, syms[l]]+Copy(xps[k], pp+1, MaxInt);
          Dec(l);
        until pp=0;
        ps:=Copy(ps, j*(1-k)+k{j, 1}, j*k+(1-k)*NoPos{high, j});
      end;
      if firstrep then
        if Pos('<mark></mark>', xps[1])=1 then begin
          xps[1]:=Copy(xps[1], Length('<mark>')+1, MaxInt);  xps[-1]:='<mark>';
        end else if (Copy(xps[1], 1, Length('</mark>'))='</mark>') and (LastDelimiter('{', wrd)>1) then begin
          xps[1]:=Copy(xps[1], Length('</mark>')+1, MaxInt);  xps[-1]:='</mark>';
        end else xps[-1]:='';
      if (ps='') and (pctatend>0) and (Pos(punctArrayNew[pctatend-1], xps[1])>0) then begin
        xps[0]:=xps[0]+xps[1];  xps[1]:='';
      end;
      Result:=xps[1]+Result;
      if ncases>2 then begin
        if ps='' then Result:=Pronoun(2, 0)+Result else if (ps[1]='(') and (ps[Length(ps)]=')') then Result:=Copy(ps, 2, Length(ps)-2)+Result else if ps<>'|' then begin
          pp:=Pos(ps, 'wvzcjfqshx')-1;
          if (Length(ps)=1) and (pp>=0) then Result:=Pronoun(pp div 5 +1, pp mod 5 +1)+Result else begin
            numer:=0;  rdiff:=0;  k:=0;  comma:=-1;  recur:=-1;
            for j:=1 to Length(ps) do case ps[j] of
              '0'..'9', 'A'..'F': try
                numer:=16*numer+HexToInt(ps[j]);
                if recur=-1 then rdiff:=16*rdiff+HexToInt(ps[j]);
                Inc(k);
              except err:=12 end;
              '#': if recur=-1 then recur:=k else err:=12;
              #30: if comma=-1 then comma:=k else err:=12;
              '-', '~', '^': ;
              '_': if j>1 then err:=12;
              else err:=1;
            end;
            if comma=-1 then comma:=k;
            if xps[1]<>punctArrayNew[9] then begin
              try
                if recur=-1 then Result:=FormatFloat(',#################0.##################', numer/IntPower(16, k-comma))+Result else begin
                  if recur>=comma then up:=1 else up:=IntPower(16, comma-recur);
                  Result:=Recurring(up*(numer-rdiff), up*(IntPower(16, k-comma)-IntPower(16, recur-comma)), err)+Result;
                end;
              except err:=12 end;
            end else if recur=-1 then begin
              try Result:=FormatFloat(',#################0.'+StringOfChar('#', Floor((k-comma)*Log10(16))), numer/IntPower(16, k-comma))+xps[0]+Result except err:=12 end;
              xps[0]:='';
            end else err:=12;
            if ps[1]='_' then Result:='\u8722?'+Result;
          end;
        end;
      end else if (ps<>'') and (ps<>'|') then err:=13+Ord(Length(ps)>1);
      Result:=xps[0]+Result;
      if i>0 then begin
        if firstrep then Result:=xps[-1]+IfThen(emact, '<mark>')+'<sup>'+IntToStr(levels[High(levels)])+IfThen(act, ActAbbr)+'</sup>'+IfThen(emact, '</mark>')+Result;
        Result:=IfThen(Copy(wrd, i-1, 1)<>'|', '-', '/')+IfThen(r>0, IfThen(rb>0, punctArrayNew[2*rb+9])+CaseAbbrs[r+31]+IfThen(rb>0, punctArrayNew[2*rb+10]))
          +IfThen(qb>0, CaseAbbrs[p mod 8]+'/')+CaseAbbrs[p mod 8 + 8*q]+IfThen((ncases>2) and (ps<>'|'), '-')+Result;
        if ncases<=2 then a[ncases]:=p div 8 else if p div 8 > 0 then err:=5;
      end;
      wrd:=Copy(wrd, 1, i-1);  xwrd:=Copy(xwrd, 1, i-1);
      firstrep:=False;
    until i=0;
    if isword {or (Pos('\u8230?', Result)>0){//} then begin     // emdashes (8230 = …)
      SetLength(levels, Length(levels)+1);
      SetLength(mainpred, Length(mainpred)+1);
      mainpred[High(mainpred)]:=pct=2;
      {if not isword then begin      //
        ncases:=3;  a[1]:=0;  a[2]:=1;
      end; {}
      if ncases>2 then begin
        if pct=-1 then err:=11 else if pct<2 then levels[High(levels)]:=levels[High(levels)-1]-(Max(a[1]+a[2]-1-2*Ord(a[2]>0)+4*pct, -1)) else if (a[1]=0) and (a[2]=1) then levels[High(levels)]:=1 else
          if (a[1]>0) and (a[2]>0) then err:=4 else if (a[1]=0) and (a[2]=0) then err:=3 else err:=6;
        act:=(a[1]=0) and (a[2]=2) and (pct=0);
        if err=0 then if levels[High(levels)]<1 then err:=8 else if (a[1]=0) and (a[2]=0) then err:=3 else if (a[1]>0) and (a[2]>0) then err:=4;
      end else err:=2;
    end else if (err=0) and emact then err:=7;
    emact:=femact;
  end; {GlossWord}
const errors: array[-1..15] of string = ({-1}'predicate skipping should be avoided', '', {1}'unglossed or otherwise invalid stem', {2}'word with too few vowels', {3}'word without accent', {4}'word with two accents',
       {5}'accent on epenthetic case', {6}'wrong accent on last word in sentence', {7}'no (non)agent to highlight', {8}'word with non-positive level', {9}'pronoun points to negative level',
       {10}'no target word of pronoun found', {11}'incomplete sentence', {12}'invalid number', {13}'invalid character after case ending', {14}'invalid characters after case ending', {15}'recurring part of number too long');
      pars = '()[]{}';
      numSymbols = ['0'..'9', 'A'..'F', '#', '_'];
      seps:    array[0..4] of string = ('--', '\u8211?-', '-\u8211?', '\u8212?-', '-\u8212?');
      sepsNew: array[0..4] of string = ('-',  '\u8211?',   '\u8211?', '\u8212?',   '\u8212?');
var i, p, q: Integer;
    st: string;
begin
  S:=Trim(S);
  if Copy(S, 1, 4)='<id=' then S:=TrimLeft(Copy(S, Pos('>', S)+1, MaxInt));
  st:=#28+S;  Result:='';  act:=False;  emact:=False;  err:=0;
  SetLength(levels, 1);    levels[0]:=1;
  SetLength(mainpred, 1);  mainpred[0]:=True;
  for i:=1 to Length(pars) do st:=StringReplace(st, '\'+pars[i], Char(i), [rfReplaceAll]);
  //if (Length(st)>0) and (st[Length(st)]='—') then st:=st+'.'; //  emdashes
  repeat
    p:=Pos('[', st);
    if p>0 then q:=Pos(']', Copy(st, p+1, MaxInt)) else q:=0;
    if q>0 then st:=Copy(st, 1, p-1)+Copy(st, p+q+1, MaxInt);
  until q=0;
  for i:=1 to Length(st) do if (st[i]=',') and (((i>1) and (st[i-1] in numSymbols)) or ((i<Length(st)) and (st[i+1] in numSymbols))) then st[i]:=#30;
  repeat
    p:=MultPos(' ,.;:', st);
    if (p<NoPos) and (Copy(st, p, 2)='  ') then Inc(p);
    Result:=Result+GlossWord(Copy(st, 1, p));
    st:=Copy(st, p+1, Maxint);
  until p=NoPos;
  for i:=0 to High(seps) do Result:=StringReplace(Result, seps[i], sepsNew[i], [rfReplaceAll]);
  for i:=2 to 4 do Result:=StringReplace(Result, ' '+punctArrayNew[2*i], punctArrayNew[2*i], [rfReplaceAll]);
  Result:=IncludeAbbrEtc(Trim(Result), ErrMsg);
  if (err=0) and (levels[High(levels)]<>1) and (S[Length(S)]<>'—') then err:=11;
  if err<>0 then ErrMsg:=ErrMsg+'Glossing '+IfThen(err>0, 'error', 'information')+' #'+IntToStr(Abs(err))+' in "'+S+'": '+errors[err]+'.'#13#10#13#10;
end; {Glossing}

procedure TTableMaker.CopyBtnClick(Sender: TObject);
var t, i, j, k, p, q, r, rs, ct, number: Integer;
    b, b2, whole: Boolean;
    st, id_, c, cls, cell, tdid, invalidId, gloss, errmsg: string;
    HTMLsN: array of array[0..3] of string;
    cs: array[0..3, 0..1] of Integer;
    empty: array[0..3] of Boolean;
    tm: TDateTime;
begin
  tm:=Now;
  errmsg:='';
  t:=TabControl.TabIndex;
  with Tables[t, ListBox.ItemIndex] do begin
    repeat
      b:=StringGrid.RowCount>2;
      for j:=0 to ColCounts[t]-1 do if HTMLs[StringGrid.RowCount-2, j]<>'' then b:=False;
      if b then begin
        StringGrid.RowCount:=StringGrid.RowCount-1;
        SetLength(HTMLs, Length(HTMLs)-1);
        SetLength(NoNutshell, Length(HTMLs));
      end;
    until not b;
    whole:=(Sender=CopyBtn) or (Sender=Copytable1);
    id_:=LowerCase(StringReplace(Id, ' ', '_', [rfReplaceAll]));
    st:='<table class="x '+LowerCase(Copy(Heads[t], 1, 3))+'"><!-- '+Heads[t]+IfThen(not whole, ' Nut')+': '+TrimLeft(Id+' ')+'-->'#13#10;
    SetLength(HTMLsN, Length(HTMLs));
    p:=0;
    for i:=0 to Length(HTMLs)-1 do if whole or not NoNutshell[i] then begin
      for j:=0 to ColCounts[t]-1 do begin
        HTMLsN[p, j]:=HTMLs[i, j];
        for b:=False to True do repeat
          q:=Pos(NutTags[0, b], HTMLsN[p, j]);
          if q>0 then begin
            r:=Pos(NutTags[1, b], Copy(HTMLsN[p, j], q, MaxInt));
            if r=0 then r:=MaxInt;
            if whole=b then HTMLsN[p, j]:=Copy(HTMLsN[p, j], 1, q-1)+Copy(HTMLsN[p, j], q+r+Length(NutTags[0, b]), MaxInt)
              else HTMLsN[p, j]:=Copy(HTMLsN[p, j], 1, q-1)+Copy(HTMLsN[p, j], q+Length(NutTags[0, b]), r-1-Length(NutTags[0, b]))+Copy(HTMLsN[p, j], q+r+Length(NutTags[0, b]), MaxInt);
          end;
        until q=0;
      end;
      Inc(p);
    end;
    SetLength(HTMLsN, p);
    for j:=0 to ColCounts[t]-1 do if HTML5Image.Visible then begin
      empty[j]:=True;
      for i:=0 to Length(HTMLsN)-1 do if HTMLsN[i, j]<>'' then empty[j]:=False;
    end else empty[j]:=False;
    number:=0;
    for i:=0 to Length(HTMLsN)-1 do if (HTMLsN[i, ColCounts[t]-1]<>'') and ((Heads[t]<>'Examples') or (Copy(HTMLsN[i, 1], Length(HTMLsN[i, 1]), 1)<>'>')) then Inc(number);
    invalidId:='';
    for i:=1 to Length(Id) do if not (Id[i] in ['0'..'9', 'A'..'Z', 'a'..'z']) then invalidId:=invalidId+Id[i];
    for i:=0 to Length(HTMLsN)-1 do begin
      j:=Ord(ColType[t, 0]=ctHead);
      if ((HTMLsN[i, 0]='') or (HTMLsN[i, 0]='>')) and (HTMLsN[i, j]='') then HTMLsN[i, j]:=' ';
      b:=False;
      for j:=0 to ColCounts[t]-1 do if Pos('xxx', HTMLsN[i, j])>0 then b:=True;
      b2:=(ColType[t,0]=ctHead) and (HTMLsN[i,0]='>');
      st:=st+'<tr'+IfThen(b or b2, ' class="'+Trim(IfThen(b, 'todo ')+IfThen(b2, 'sep'))+'"')+'>';
      if (i=0) or ((ColType[t,0]=ctHead) and (HTMLsN[i,0]<>'')) then begin
        rs:=0;  j:=0;
        while (rs=0) or (i+rs<Length(HTMLsN)) and ((ColType[t, 0]<>ctHead) or (HTMLsN[i+rs, 0]='')) do begin
          if InterGlossing and (Trim(HTMLsN[i+rs, 1])<>'') and (HTMLsN[i+rs, 1][1]<>' ') and not ((HTMLsN[i+rs, 1][1]='(') and (HTMLsN[i+rs, 1][Length(HTMLsN[i+rs, 1])]=')')) then Inc(j);
          Inc(rs);
        end;
        st:=st+'<th'+IfThen(rs+j>1, ' rowspan="'+IntToStr(rs+j)+'"');
        if i=0 then st:=st+' title="'+Copy(Heads[t], 1, Length(Heads[t])-Ord((Heads[t]<>'Vocabulary') and (number<=1)))+'."';
        st:=st+'></th>'#13#10;                                                                         
      end;
      for k:=0 to 1 do begin
        cs[ColCounts[t]-1, k]:=1;
        for j:=ColCounts[t]-2 downto 0 do if HTMLsN[i, j+1]='' then cs[j, k]:=cs[j+1, k]+Ord(not empty[j+1] or (k=1)) else cs[j, k]:=1;
        if ColType[t,0]=ctHead then Dec(cs[0, k]);
      end;
      gloss:='';
      for j:=0 to ColCounts[t]-1 do if ((HTMLsN[i, j]<>'')) and not ((ColType[t, j]=ctHead) and (HTMLsN[i, j]='>')) then begin
        if (ColType[t, j]=ctLem) and (HTMLsN[i, j][1]='(') and (HTMLsN[i, j][Length(HTMLsN[i, j])]=')') then begin
          ct:=ctEng;           cell:=Copy(HTMLsN[i, j], 2, Length(HTMLsN[i, j])-2);
        end else begin
          ct:=ColType[t, j];   cell:=HTMLsN[i, j];
        end;
        cell:=SubstituteStems(cell, ct=ctLem, Options.Dictionary.Items);
        tdid:='';
        if Copy(cell, 1, 4)='<id=' then begin
          p:=Pos('>', cell);
          if p>0 then begin
            tdid:=' id="'+StringReplace(Copy(cell, 5, p-5), '"', '', [rfReplaceAll])+'"';
            cell:=Copy(cell, p+1, MaxInt);
          end;
        end;
        if ColType[t, j]<>ctHead then c:='d' else begin
          b:=False;
          for k:=j+1 to ColCounts[t]-1 do if HTMLsN[i, k]<>'' then b:=True;
          c:='h scope="'+IfThen(b, 'row', 'col')+'"';
        end;
        if cs[j, 0]>1 then c:=c+' colspan="'+IntToStr(cs[j, 0])+'"';
        if ct=ctLit then cls:='literal' else cls:='';
        if (cs[j, 1]>1) and HTML5Image.Visible then cls:=cls+' '+'span'+IntToStr(cs[j, 1]);
        if cls<>'' then c:=c+' class="'+Trim(cls)+'"';
        if ct=ctLem then c:=c+' lang="x-lm"'+LemTranscript(cell, Symbols, errmsg);
        if ColType[t, j]<>ctSolv then begin
          st:=st+'<t'+c+tdid+'>';
          if ct=ctLem then begin
            if Heads[t]='Vocabulary' then st:=st+'<a href="../lemeng/index.php?lemma='+cell+'">';
            st:=st+Trim(IncludeBracketsEtc(cell, True));
            if Heads[t]='Vocabulary' then st:=st+'</a>';
          end else st:=st+IncludeAbbrEtc(Trim(cell), errmsg);
        end else begin
          st:=st+'<t'+c+' id="u'+IntToStr(Un)+'_'+id_+'_'+cell+'"><a href="solutions.php?u='+IntToStr(Un)+'&amp;e='+id_+'&amp;i='+cell
            +'" onClick="ajaxSolve()">'
            +'<img src="../images/solve.svg" alt="Solve"></a>';
          if invalidId<>'' then errmsg:=errmsg+'Table error: invalid characters "'+invalidId+'" in table id'#13#10#13#10;
          invalidId:='';
        end;
        st:=st+'</t'+c[1]+'>';
        if InterGlossing then
          if (ct=ctLem) and (Trim(cell)<>'') and (cell[1]<>' ') then begin
            p:=0;
            for k:=j to ColCounts[t]-1 do if not empty[k] then Inc(p);
            gloss:='<tr><td colspan="'+IntToStr(p)+'" class="igloss">'+Glossing(HTMLsN[i,j], errmsg)+'</td></tr>'#13#10;
          end else TestLemParts(HTMLsN[i,j], ct=ctLem, errmsg);
      end;
      st:=st+'</tr>'#13#10+gloss;
    end;
    j:=-1;
    if ColType[t,0]=ctHead then for i:=0 to Length(HTMLsN)-1 do if (HTMLsN[i,0]<>'') and (HTMLsN[i,0]<>'>') and (HTMLsN[i,1]<>'') then j:=i;
    if j>-1 then errmsg:=errmsg+'Table error in line '+IntToStr(j+1)+': header lines cannot contain an entry in the second column'#13#10#13#10;
  end;
  Clipboard.Open;
  Clipboard.Clear;
  CopyRTFasUnicode(HTML4or5(st)+'</table>'#13#10);
  Clipboard.Close;
  TimeLabel.Caption:='Copied '+IfThen(not whole, 'for nutshells ')+'in '+FloatToStrF((Now-tm)*86400, ffGeneral, 4, 2)+' s';
  ErrorMessageBox(errmsg);
  if SolvPanel.Visible then SolvMemo.SetFocus else StringGrid.SetFocus;
end; {CopyBtnClick}

procedure TTableMaker.StringGridSelectCell(Sender: TObject; ACol, ARow: Integer; var CanSelect: Boolean);
begin
  EnableLemBar(ACol);
end; {StringGridSelectCell}

procedure TTableMaker.ControlEnterExit(Sender: TObject);
begin
  EnableLemBar(StringGrid.Col);
end; {ControlEnterExit}

procedure TTableMaker.StringGridKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Shift=[]) and (Key=VK_Down) and (StringGrid.Row=StringGrid.RowCount-1) then NewLineBtnClick(nil);
  if Shift=[ssShift, ssCtrl] then case Key of
    88{X}: Cutlines1Click(Cutlines1);
    67{C}: Copylines1Click(Copylines1);
    86{V}: if Clipboard.HasFormat(CF_Rows[TabControl.TabIndex]) then Pastelines1Click(Pastelines1);
    65{A}: Selectall1Click(Selectall1);
  end;
  if Shift=[ssCtrl] then case Key of
    70{F}: FocusControl(FindEdit);
    $37{{}: ToolHilightClick(ToolHilight);
    $38{[}: ToolHilightClick(ToolLevel);
    $39{(}: ToolHilightClick(ToolStem);
    $31{<lem>}: ToolHilightClick(ToolLem);
    $32{<nut>}: ToolHilightClick(ToolNut);
    $33{<nonut>}: ToolHilightClick(ToolNonut);
  end;
  if (Shift=[ssShift]) and (Key=VK_Delete) then Deletelines1Click(Deletelines1);
  if (Shift=[]) and (Key=VK_F3) then
    if FindEdit.Text='' then FocusControl(FindEdit) else FindEditKeyDown(Sender, Key, Shift);
  if ((Shift=[]) and (Key=VK_F4))
      or (MultiSelectBtn.Down and (((Key in [VK_Space, 48..57, 65..90, 96..111]) and (Shift-[ssShift]=[])) or ((Key in [VK_Left, VK_Right, VK_F2]) and (Shift=[]))))
      or (not MultiSelectBtn.Down and not StringGrid.EditorMode and (Key in [VK_Up, VK_Down, VK_Left, VK_Right]) and (Shift=[ssShift])) then begin
    MultiSelectBtn.Down:=not MultiSelectBtn.Down;
    MultiSelectBtnClick(nil);
  end;
  if (Shift=[]) and (Key=VK_F10) then NeogrammBtnClick(nil);
  if (Shift=[]) and (Key=VK_F5) then NutshellBtnClick(nil);
  if Key=VK_F9 then
    if Shift=[]        then ShowOpts(0) else
    if Shift=[ssShift] then ShowOpts(1) else
    if Shift=[ssCtrl]  then ShowOpts(2) else
    if Shift=[ssAlt]   then ShowOpts(3) else
    if Shift=[ssShift, ssCtrl] then ShowOpts(4) else
    if Shift=[ssShift, ssAlt] then ShowOpts(5);
end; {StringGridKeyDown}

procedure TTableMaker.StringGridGetEditText(Sender: TObject; ACol, ARow: Integer; var Value: string);
begin
  if (ARow>0) and (ColType[TabControl.TabIndex, ACol]=ctLem) then SetLemFont(StringGrid.Font, True, 0) else StringGrid.ParentFont:=True;
end; {StringGridGetEditText}

function TTableMaker.GetFreeSolutionNo(ACol, ARow: Integer): string;
var b: Boolean;
    i, r: Integer;
begin
  r:=0;
  with ThisTable do repeat
    b:=True;
    Inc(r);
    for i:=0 to Length(HTMLs)-1 do if (i<>ARow-1) and (HTMLs[i, ACol]=IntToStr(r)) then b:=False;
  until b;
  Result:=IntToStr(r);
end; {GetFreeSolutionNo}

procedure TTableMaker.StringGridSetEditText(Sender: TObject; ACol, ARow: Integer; const Value: string);
begin
  with ThisTable do if ColType[TabControl.TabIndex, ACol]=ctSolv then begin
    if Value='' then HTMLs[ARow-1, ACol]:='' else begin
      HTMLs[ARow-1, ACol]:=GetFreeSolutionNo(ACol, ARow);
      StringGrid.Cells[ACol, ARow]:=HTMLs[ARow-1, ACol];
    end;
  end else HTMLs[ARow-1, ACol]:=TrimRight(Value);
  StringGrid.Invalidate;
end; {StringGridSetEditText}

procedure TTableMaker.StringGridDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect; State: TGridDrawState);
var t: TTable;
  procedure DrawString(S, F: string; P: Integer; var X: Integer);
  var b: Boolean;
      q: array[False..True] of Integer;
      st: string;
  begin
    if S<>'' then with StringGrid.Canvas do begin
      Font.Name:=F;
      if F=LemFont then Inc(P, 2) else if F<>TableMaker.Font.Name then Inc(P, 1);
      Font.Size:=TableMaker.Font.Size+P;
      for b:=False to True do q[b]:=MaxInt; 
      repeat
        with Brush do for b:=Color=NutColors[True, t.Nut<255] to (Color=NutColors[True, t.Nut<255]) or (Style=bsClear) do begin
          q[b]:=Pos(NutTags[Ord(Style<>bsClear), b], S);
          if q[b]=0 then q[b]:=Length(S)+1;
        end;
        b:=q[True]<q[False];
        st:=Copy(S, 1, q[b]-1);
        TextOut(X, Rect.Top+2*Ord(P>=0)-LemFontCorr*Ord(F=LemFont), st);
        Inc(X, TextWidth(st));
        if q[b]<Length(S)+1 then
          if Brush.Style=bsClear then Brush.Color:=NutColors[b, t.Nut<255] else Brush.Style:=bsClear;
        S:=Copy(S, q[b]+Length(NutTags[Ord(Brush.Style=bsClear), b]), MaxInt);
      until Length(S)=0;
    end;
  end; {DrawString}
  procedure DrawLemString(S: string; var X: Integer);
  var p, q, b, i: Integer;
      esc: Boolean;
      cl: TColor;
  begin
    esc:=False;
    with StringGrid.Canvas do repeat
      p:=MultPos('([{}\', S);
      if (p=1) and esc then p:=MultPos('([{}\', Copy(S, 2, 1000000000))+1;
      esc:=False;
      p:=Min(MultPos(Symbols[tsyCode], S), p);
      if LemBox.Checked then DrawString(Copy(S, 1, p-1), LemFont, 0, X) else DrawString(Copy(S, 1, p-1), SansSerif, 0, X);
      if p<=Length(S) then case S[p] of
        '(', '[': begin
          q:=Pos(Char(41+52*Ord(S[p]='[')), S);
          if q=0 then q:=Length(S)+1;
          b:=Ord((S[p]='(') and not LemBox.Checked);
          if Brush.Style=bsClear then cl:=clNone else cl:=Brush.Color;
          if SubstStems and (S[p]='(') and (Options.Dictionary.Items.IndexOfName(Copy(S, p+1, q-p-1))=-1) then
            Brush.Color:=IfThen((Font.Color=clWindowText) or (Font.Color=clGreen), $c0c0ff, $606080);
          DrawString(Copy(S, p+1-b, q-p-1+2*b), Serif, -2*Ord(S[p]='['), X);
          if cl=clNone then Brush.Style:=bsClear else Brush.Color:=cl;
          p:=q;
        end;
        '{': if Font.Color=clWindowText then Font.Color:=clGreen else if Font.Color=clHighlightText then Font.Color:=clLime;
        '}': if Font.Color=clGreen then Font.Color:=clWindowText else if Font.Color=clLime then Font.Color:=clHighlightText;
        '\': esc:=True;
        else for i:=0 to Length(Symbols[0])-1 do if (Symbols[tsyCode, i]<>'') and (Symbols[tsyCode, i]=Copy(S, p, Length(Symbols[tsyCode, i]))) then begin
          DrawString(Symbols[tsyCode, i], SansSerif, 0, X);
          Inc(p, Length(Symbols[tsyCode, i])-1);
          Break;
        end;
      end;
      S:=Copy(S, p+1, 1000000000);
    until Length(S)=0;
  end; {DrawLemString}
var ct, p, q, x: Integer;
    st: string;
begin
  t:=ThisTable;
  ct:=ColType[TabControl.TabIndex, ACol];
  st:=StringGrid.Cells[ACol, ARow];
  x:=Rect.Left+2;
  if t<>nil then with StringGrid.Canvas do begin
    FillRect(Classes.Rect(Rect.Left, Rect.Top, Rect.Right, Rect.Bottom));
    Brush.Style:=bsClear;
    if Copy(st, 1, 4)='<id=' then begin
      p:=Pos('>', st);
      if p>0 then begin
        DrawString('<id>', Serif, -2, x);
        st:=Copy(st, p+1, MaxInt);
      end;
    end;
    if (ARow=0) or (ct=ctSolv) then begin
      if (ct=ctSolv) and (ARow>0) then
        if t.HTMLs[ARow-1, ACol+1]='' then Font.Color:=clSilver else if t.HTMLs[ARow-1, ACol]='' then ImageList.Draw(StringGrid.Canvas, Rect.Left, Rect.Top, 15);
      DrawString(st, TableMaker.Font.Name, 0, x);
    end else if (ct<>ctLem) or ((st<>'') and (st[1]='(') and (st[Length(st)]=')')) then begin
      case ct of
        ctHead: Font.Style:=[fsBold];
        ctLit: Font.Style:=[fsItalic];
      end;
      if (ct=ctLem) and (ARow>0) then st:=Copy(st, 2, Length(st)-2);
      repeat
        p:=Pos('<lem>', st);
        if p=0 then p:=Length(st)+1;
        DrawString(Copy(st, 1, p-1), Serif, 0, x);
        if p<=Length(st) then begin
          q:=Pos('</lem>', st);
          if q=0 then q:=Length(st)+1;
          DrawLemString(Copy(st, p+5, q-p-5), x);
        end else q:=p;
        st:=Copy(st, q+6, MaxInt);
      until Length(st)=0;
    end else DrawLemString(st, x);
    if (ARow>0) and (ACol=ColCounts[TabControl.TabIndex]-1) and t.NoNutshell[ARow-1] then
      ImageList.Draw(StringGrid.Canvas, Rect.Right-16, (Rect.Top+Rect.Bottom) div 2 -8, 18+Ord(t.Nut=255));
  end;
end; {StringGridDrawCell}

procedure TTableMaker.StringGridMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  StringGrid.MouseToCell(X, Y, MouseC, MouseR);
end;

procedure TTableMaker.StringGridClick(Sender: TObject);
begin
  UpdateTableButtons;
end; {StringGridClick}

procedure TTableMaker.StringGridDblClick(Sender: TObject);
var p: TPoint;
    c, r: Integer;
    sol: string;
begin
  p:=StringGrid.ScreenToClient(Mouse.CursorPos);
  StringGrid.MouseToCell(p.x, p.y, p.x, p.y);
  if ColType[TabControl.TabIndex, p.x]=ctSolv then begin
    sol:=FullSolutionsPath;
    SolvX:=p.x+1;  SolvY:=p.y-1;
    if p.y=0 then Sortsolutions1Click(nil) else Editsolution1Click(nil);
  end else with StringGrid do if not EditorMode then begin
    MultiSelectBtn.Down:=False;
    MultiSelectBtnClick(nil);
    p:=ScreenToClient(Mouse.CursorPos);
    MouseToCell(p.X, p.Y, c, r);
    if (c>-1) and (r>0) then begin
      Col:=c;  Row:=r;
      EditorMode:=True;
    end;
  end;
end; {StringGridDblClick}

procedure TTableMaker.SolvMemoChange(Sender: TObject);
begin
  ThisTable.HTMLs[SolvY, SolvX]:=SolvMemo.Text;
end; {SolvMemoChange}

procedure TTableMaker.SolvOKBtnClick(Sender: TObject);
begin
  SolvPanel.Hide;
  StringGrid.Enabled:=True;
end; {SolvOKBtnClick}

procedure TTableMaker.Edit1Click(Sender: TObject);
begin
  with MultiSelectBtn do if Down then begin
    StringGrid.Row:=MouseR;
    Down:=False;
    Click;
    StringGrid.Col:=MouseC;
  end;
  StringGrid.EditorMode:=True;
end; {Edit1Click}

procedure TTableMaker.Editsolution1Click(Sender: TObject);
begin
  if StringGrid.Cells[SolvX-1, SolvY+1]='' then begin
    StringGrid.Cells[SolvX-1, SolvY+1]:=GetFreeSolutionNo(SolvX-1, SolvY+1);
    ThisTable.HTMLs[SolvY, SolvX-1]:=StringGrid.Cells[SolvX-1, SolvY+1];
  end;
  SolvCaption.Caption:='Solution file '+IntToStr(UnitCombo.ItemIndex)+'_'+LowerCase(IdEdit.Text)+'_'+StringGrid.Cells[SolvX-1, SolvY+1]+'.html for: '
    +StringGrid.Cells[SolvX-3, SolvY+1]+' '+StringGrid.Cells[SolvX-2, SolvY+1];
  SolvMemo.Text:=ThisTable.HTMLs[SolvY, SolvX];
  SolvPanel.Show;
  EnableLemBar(StringGrid.Col);
  ActiveControl:=SolvMemo;
  StringGrid.Enabled:=False;
end; {Editsolution1Click}

procedure TTableMaker.Sortsolutions1Click(Sender: TObject);
var n, i, c: Integer;
begin
  with ThisTable do begin
    GetNoOfSolutions(1, Length(HTMLs), UnitCombo.ItemIndex, IdEdit.Text, n, c);
    if n>1 then begin
      n:=1;
      for i:=0 to Length(HTMLs)-1 do if HTMLs[i, SolvX-1]<>'' then begin
        HTMLs[i, SolvX-1]:=IntToStr(n);
        StringGrid.Cells[SolvX-1, i+1]:=IntToStr(n);
        Inc(n);
      end;
    end;
  end;
end; {Sortsolutions1Click}

function TTableMaker.NewLine(Strings: array of string; Line: Integer): Integer;
var i, j: Integer;
begin
  Result:=-1;
  with StringGrid, ThisTable do begin
    SetFocus;
    Col:=0;
    RowCount:=RowCount+1;
    SetLength(HTMLs, Length(HTMLs)+1);
    SetLength(NoNutshell, Length(HTMLs));
    for i:=RowCount-2 downto Line do begin
      for j:=0 to ColCount-1+Ord(TabControl.TabIndex=SolvTab) do begin
        HTMLs[i, j]:=HTMLs[i-1, j];
        if j<ColCount then Cells[j, i+1]:=Cells[j, i];
      end;
      NoNutshell[i]:=NoNutshell[i-1];
    end;
    NoNutshell[Line-1]:=False;
    for i:=0 to ColCount-1+Ord(TabControl.TabIndex=SolvTab) do begin
      if (ColType[TabControl.TabIndex, i]<>ctSolv) or (Strings[i]='') then HTMLs[Line-1, i]:=Strings[i] else begin
        HTMLs[Line-1, i]:=GetFreeSolutionNo(i, Line);
        if HTMLs[Line-1, i]<>'' then HTMLs[Line-1, i+1]:=Strings[i] else Result:=i;
      end;
      Cells[i, Line]:=HTMLs[Line-1, i];
    end;
  end;
end; {NewLine}

procedure TTableMaker.NewLineBtnClick(Sender: TObject);
begin
  SolvOKBtnClick(nil);
  NewLine(['', '', '', ''], StringGrid.Row+1);
  StringGrid.Row:=StringGrid.Row+1;
  StringGrid.EditorMode:=True;
end; {NewLineBtnClick}

procedure TTableMaker.LinesBtnClick(Sender: TObject);
var P: TPoint;
begin
  SolvOKBtnClick(nil);
  Pastelines1.Enabled:=Clipboard.HasFormat(CF_Rows[TabControl.TabIndex]);
  P:=GridPanel.ClientToScreen(Point(LinesBtn.Left, LinesBtn.Top+LinesBtn.Height));
  ClipboardMenu.Popup(P.X, P.Y);
  if Visible then StringGrid.SetFocus;
end; {LinesBtnClick}

procedure TTableMaker.Cutlines1Click(Sender: TObject);
begin
  Copylines1Click(nil);
  Deletelines1Click(nil);
end; {Cutlines1Click}

procedure TTableMaker.Copylines1Click(Sender: TObject);
var i, j: Integer;
    st, ins, sol: string;
    H: THandle;
    P: PChar;
begin
  st:='';
  sol:=FullSolutionsPath;
  with StringGrid do for i:=Selection.Top to Selection.Bottom do begin
    if ThisTable.NoNutshell[i-1] then st:=st+'¤';
    for j:=0 to ColCount-1 do begin
      if ColType[TabControl.TabIndex, j]<>ctSolv then ins:=Cells[j, i] else ins:=ThisTable.HTMLs[i-1, j+1];
      st:=st+StringReplace(StringReplace(ins, '¤', '\u164?', [rfReplaceAll]), '|', '\u124?', [rfReplaceAll])+'|';
    end;
  end;
  H:=GlobalAlloc(gmem_Moveable, Length(st)+1);
  P:=GlobalLock(H);
  StrPCopy(P, st);
  GlobalUnlock(H);
  Clipboard.SetAsHandle(CF_Rows[TabControl.TabIndex], H);
end; {Copylines1Click}

procedure TTableMaker.Pastelines1Click(Sender: TObject);
var i, p, q, r, delSolv: Integer;
    n: Boolean;
    st: string;
    H: THandle;
    Ptr: PChar;
    strings: array[0..3] of string;
begin
  H:=Clipboard.GetAsHandle(CF_Rows[TabControl.TabIndex]);
  Ptr:=GlobalLock(H);
  st:=StrPas(Ptr);
  GlobalUnlock(H);
  p:=0;   delSolv:=-1;
  r:=StringGrid.Row;
  with StringGrid, ThisTable do while p<Length(st) do begin
    for i:=0 to ColCount-1 do begin
      q:=Pos('|', Copy(st, p+1, MaxInt));
      strings[i]:=StringReplace(Copy(st, p+1, q-1), '\u124?', '|', [rfReplaceAll]);
      Inc(p, q);
    end;
    n:=Copy(strings[0], 1, 1)='¤';
    if n then strings[0]:=Copy(strings[0], 2, MaxInt);
    if delSolv>-1 then strings[delSolv]:='';
    delSolv:=Max(delSolv, NewLine(strings, r));   
    if n then NoNutshell[r-1]:=True;
    Inc(r);
  end;
  StringGrid.Row:=r-1;
  UpdateTableButtons;
end; {Pastelines1Click}

procedure TTableMaker.Deletelines1Click(Sender: TObject);
var i, j, n: Integer;
begin
  with StringGrid, ThisTable do begin
    n:=Selection.Bottom-Selection.Top+1;
    if (n=1) or (Sender=nil) or (TangoMessageBox('Delete '+IntToStr(n)+' lines?', mtConfirmation, [mbYes, mbNo], '')=idYes) then begin
      with StringGrid do if RowCount=Selection.Bottom-Selection.Top+2 then begin
        NewLine(['', '', '', ''], RowCount);
        Selection:=TGridRect(Rect(0, 1, 1, RowCount-2));
      end;
      for i:=Selection.Top to RowCount-n-1 do begin
        for j:=0 to ColCount-1+Ord(TabControl.TabIndex=SolvTab) do begin
          HTMLs[i-1, j]:=HTMLs[i+n-1, j];
          if j<ColCount then Cells[j, i]:=Cells[j, i+n];
        end;
        NoNutshell[i-1]:=NoNutshell[i+n-1];
      end;
      RowCount:=RowCount-n;
      SetLength(HTMLs, Length(HTMLs)-n);
      SetLength(NoNutshell, Length(HTMLs));
    end;
    UpdateTableButtons;
  end;
end; {Deletelines1Click}

procedure TTableMaker.Selectall1Click(Sender: TObject);
begin
  if not MultiSelectBtn.Down then MultiSelectBtnClick(Selectmlines1);
  StringGrid.Selection:=TGridRect(Rect(0, 1, StringGrid.ColCount-1, StringGrid.RowCount-1));
end; {Selectall1Click}

procedure TTableMaker.NutshellBtnClick(Sender: TObject);
var i: Integer;
begin
  SolvOKBtnClick(nil);
  if (Notebook.PageIndex=0) and (GridPanel.Visible) then begin
    if Sender<>NutshellBtn then NutshellBtn.Down:=not NutshellBtn.Down;
    with StringGrid.Selection do for i:=Top to Bottom do ThisTable.NoNutshell[i-1]:=NutshellBtn.Down;
    StringGrid.Invalidate;  
  end;
end; {NutshellBtnClick}

procedure TTableMaker.MultiSelectBtnClick(Sender: TObject);
begin
  SolvOKBtnClick(nil);
  if (Notebook.PageIndex=0) and (GridPanel.Visible) then begin
    if Sender=Selectmlines1 then MultiSelectBtn.Down:=not MultiSelectBtn.Down;
    Selectmlines1.Checked:=MultiSelectBtn.Down;
    with StringGrid do if MultiSelectBtn.Down then Options:=Options+[goRowSelect, goRangeSelect]-[goEditing]
      else Options:=Options-[goRowSelect, goRangeSelect]+[goEditing];
    if Visible then StringGrid.SetFocus;
  end;
end; {MultiSelectBtnClick}

procedure TTableMaker.UpDownLineBtnClick(Sender: TObject);
var i, j: Integer;
    aux: string;
    auxn: Boolean;
    rect: TGridRect;
begin
  SolvOKBtnClick(nil);
  with StringGrid, ThisTable do if (Sender=UpLineBtn) or (Sender=Up2) then begin
    for i:=0 to ColCount+Ord(TabControl.TabIndex=SolvTab)-1 do begin
      if i<ColCount then for j:=Selection.Top to Selection.Bottom do Cols[i].Exchange(j-1, j);
      aux:=HTMLs[Selection.Top-2, i];
      for j:=Selection.Top to Selection.Bottom do HTMLs[j-2, i]:=HTMLs[j-1, i];
      HTMLs[Selection.Bottom-1, i]:=aux;
    end;
    auxn:=NoNutshell[Selection.Top-2];
    for j:=Selection.Top to Selection.Bottom do NoNutshell[j-2]:=NoNutshell[j-1];
    NoNutshell[Selection.Bottom-1]:=auxn;
  end else begin
    for i:=0 to ColCount+Ord(TabControl.TabIndex=SolvTab)-1 do begin
      if i<ColCount then for j:=Selection.Bottom downto Selection.Top do Cols[i].Exchange(j, j+1);
      aux:=HTMLs[Selection.Bottom, i];
      for j:=Selection.Bottom downto Selection.Top do HTMLs[j, i]:=HTMLs[j-1, i];
      HTMLs[Selection.Top-1, i]:=aux;
    end;
    auxn:=NoNutshell[Selection.Bottom];
    for j:=Selection.Bottom downto Selection.Top do NoNutshell[j]:=NoNutshell[j-1];
    NoNutshell[Selection.Top-1]:=auxn;
  end;
  rect:=StringGrid.Selection;
  rect.Top:=StringGrid.Selection.Top+TComponent(Sender).Tag;
  rect.Bottom:=StringGrid.Selection.Bottom+TComponent(Sender).Tag;
  StringGrid.Selection:=rect;
  UpdateTableButtons;
  StringGrid.SetFocus;
end; {UpDownLineBtnClick}

procedure TTableMaker.UpdateTableButtons;
begin
  UpLineBtn.Enabled:=StringGrid.Selection.Top>1;
  DownLineBtn.Enabled:=StringGrid.Selection.Bottom<StringGrid.RowCount-1;
  if (ListBox.Items.Count>0) and (ListBox.ItemIndex>-1) then NutshellBtn.Down:=ThisTable.NoNutshell[StringGrid.Row-1];
end; {UpdateTableButtons}

{------------------------------------------------ Bottom -----------------------------------------------------------}

procedure TTableMaker.LemButtonClick(Sender: TObject);
begin
  with TToolButton(Sender) do if (ActiveControl=FindEdit) or (ActiveControl=SolvMemo) then TCustomEdit(ActiveControl).SelText:=Caption[1] else
    if TabControl.TabIndex=Length(Heads)-1 then Memo.SelText:=Caption[1]
      else StringGrid.Perform(WM_Char, Ord(Caption[1]), 0);
end; {LemButtonClick}

procedure TTableMaker.ToolButtonAstroClick(Sender: TObject);
var P: TPoint;
begin
  P:=LemBar.ClientToScreen(Point(ToolButtonAstro.Left, ToolButtonAstro.Top+ToolButtonAstro.Height));
  AstroMenu.Popup(P.X, P.Y);
end; {AstroButtonClick}

procedure TTableMaker.AstroMenuClick(Sender: TObject);
begin
  with TMenuItem(Sender) do if (ActiveControl=FindEdit) or (ActiveControl=SolvMemo) then TCustomEdit(ActiveControl).SelText:=Char(Tag) else
    if TabControl.TabIndex=Length(Heads)-1 then Memo.SelText:=Char(Tag)
      else StringGrid.Perform(WM_Char, Tag, 0);
end; {AstroMenuClick}

procedure TTableMaker.NeogrammBtnClick(Sender: TObject);
begin
  if not FileExists(NeogrammPath+'\Neogramm.exe') then SelectDirectory('Locate Neogrammarian directory:', '', NeogrammPath);
  TExecThread.Create(NeogrammPath+'\Neogramm.exe', SW_ShowNormal);
  if Sender=NeogrammBtn then if ActiveControl=NeogrammBtn then if Notebook.PageIndex=1 then Memo.SetFocus else if SolvPanel.Visible then SolvMemo.SetFocus else StringGrid.SetFocus;
end;

procedure TTableMaker.ToolHilightClick(Sender: TObject);
const ins: array[0..11] of string = ('{', '}', '[', ']', '(', ')', '<lem>', '</lem>', '<nut>', '</nut>', '<nonut>', '</nonut>');
var i: Integer;
begin
  with TSpeedButton(Sender) do if TabControl.TabIndex=Length(Heads)-1 then Memo.SelText:=ins[Tag]+Memo.SelText+ins[Tag+1] else
      if ActiveControl=SolvMemo then SolvMemo.SelText:=ins[Tag]+SolvMemo.SelText+ins[Tag+1] else begin
    for i:=1 to Length(ins[Tag])   do StringGrid.Perform(WM_Char, Ord(ins[Tag]  [i]), 0);
    for i:=1 to Length(ins[Tag+1]) do StringGrid.Perform(WM_Char, Ord(ins[Tag+1][i]), 0);
  end;
end; {HilightBtnClick}

procedure TTableMaker.FindClearClick(Sender: TObject);
begin
  FindEdit.Text:='';
  FindEdit.Color:=clWindow;
end; {FindClearClick}

procedure TTableMaker.FindEditKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Shift=[] then case Key of
   VK_Return, VK_F3: FindEditChange(FindDown);
   VK_Escape: FindClearClick(nil);
  end;
  if (Shift=[ssShift]) and (Key=VK_F3) then FindEditChange(FindUp);
end; {FindEditKeyDown}

procedure TTableMaker.FindEditChange(Sender: TObject);
  function TextToCompare(st: string; ColType: Integer{FindEdit: -1}; ExcludeCurly: Boolean): string;
  const vowels = 'àèÌìòÒùÙ`áéýíóÓúÚ´aeyioOuU ';
  var i, d, n: Integer;
  begin
    d:=1;
    if (ColType=ctLem) and (Copy(st, 1, 4)='<id=') then while (d<Length(st)) and (st[d+1]<>'>') do begin
      if LemFindBtn.Tag=1 then st[d]:=#1;
      Inc(d);
    end;
    if ExcludeCurly then for i:=0 to 1 do st:=StringReplace(st, Char(123+2*i), '', [rfReplaceAll]);
    n:=0;
    Result:=st;
    if ColType>-1 then for i:=d to Length(st) do begin
      if (ColType<>ctLem) and (i>5) and (Copy(st, i-5, 5)='<lem>') then begin ColType:=ctLem; Inc(n); end;
      if (ColType=ctLem) and (st[i]='(') and ((i=1) or (st[i-1]<>'\')) then begin ColType:=ctEng; if (i>1) or (st[Length(st)]<>')') then Inc(n); end;
      if (ColType=ctLem) and (Copy(st, i, 6)='</lem>') and (n>0) then begin ColType:=ctEng; Dec(n); end;
      if (ColType<>ctLem) and (i>1) and (st[i-1]=')') and ((n=2) or (st[i-2]<>'\')) and (n>0) then begin ColType:=ctLem; Dec(n); end;
      if (LemFindBtn.Tag<>2) and ((ColType=ctLem) xor (LemFindBtn.Tag=1)) then Result[i]:=#1;
    end;
    if LemFindBtn.Tag=1 then begin
      if not AccentCheckBox.Checked then for i:=0 to 17 do Result:=StringReplace(Result, vowels[i+1], vowels[i mod 9 +19], [rfReplaceAll]);
      if not VowelCheckBox.Checked  then for i:=0 to 26 do if i mod 9<>8 then Result:=StringReplace(Result, vowels[i+1], vowels[i div 9 *9 +1], [rfReplaceAll]);
    end else if not CaseCheckBox.Checked then Result:=LowerCase(Result);
  end; {TextToCompare}

var t, l, r, c: Integer;
    memotxt: string;
  procedure WrapTab(n: Integer);
  begin
    repeat t:=(t+TabControl.Tabs.Count+n) mod TabControl.Tabs.Count until (t=High(Heads)) or (Length(Tables[t])>0);
    if t=High(Heads) then begin
      if l<0 then begin
        l:=Length(TextPages)-1;  c:=Length(TextPages[l].Text)-1;
      end else begin
        l:=0;  c:=0;
      end;
    end else begin
      if l<0 then begin
        l:=Length(Tables[t])-1;  r:=Length(Tables[t,l].HTMLs)-1;  c:=ColCounts[t]-1;
      end else begin
        l:=0;  r:=0;  c:=0;
      end;
    end;
  end; {WrapTab}
  procedure Hop(n: Integer);
  begin
    if t<High(Heads) then begin
      if (r=-1) then c:=(n-1) div 2 else if (c=0) and (n=-1) then c:=-1 else c:=(c+n) mod ColCounts[t];
      if c=(n-1) div 2 then begin
        if (r=-1) and (n=-1) then r:=-2 else if Length(Tables[t])=0 then r:=(n-1) div 2 -1 else r:=(r+n+1) mod (Length(Tables[t,l].HTMLs)+1) -1;
        if r=(n-1) div 2 -1 then begin
          if (l=0) and (n=-1) then l:=-1 else if Length(Tables[t])=0 then l:=(n-1) div 2 else l:=(l+n) mod Length(Tables[t]);
          if l=(n-1) div 2 then WrapTab(n);
        end;
      end;
      if (n=-1) and (t<High(Heads)) then begin
        if l=-1 then l:=Length(Tables[t])-1;
        if r=-2 then r:=Length(Tables[t,l].HTMLs)-1;
        if c=-1 then c:=ColCounts[t]-1;
      end;
    end else begin
      Inc(c, n);                                             
      if (c<0) or (c>Length(TextPages[l].Text)) then begin
        Inc(l, n);
        memotxt:='';
        if (l<0) or (l>=Length(TextPages)) then WrapTab(n) else if c<0 then c:=Length(TextPages[l].Text) else c:=0;
      end;
    end;
  end; {Hop}
var fnd, txt: string;
    f, ft, excurly: Boolean;
    i: Integer;
    ti: TDateTime;
begin                      
  if TableMaker.Visible then FocusControl(FindEdit);
  FindClear.Enabled:=FindEdit.Text<>'';
  FindDown.Enabled:=FindClear.Enabled;
  FindUp.Enabled:=FindClear.Enabled;
  fnd:=TextToCompare(FindEdit.Text, -1, False);
  excurly:=(LemFindBtn.Tag>0) and (MultPos('{}', fnd)=NoPos);
  f:=False;  ft:=False;
  ti:=Now;
  if FindEdit.Text<>'' then begin
    t:=TabControl.TabIndex;
    if t<High(Heads) then begin
      l:=ListBox.ItemIndex;
      if FoundTitle then r:=-1 else r:=StringGrid.Row-1;
      c:=StringGrid.Col;
    end else begin
      l:=UnitList.ItemIndex;
      c:=Memo.SelStart+1-Ord((Sender=FindEdit) or (Sender is TCheckBox));
      Dec(c, Length(Copy(TextPages[l].Text, 1, c))-Length(TextToCompare(Copy(TextPages[l].Text, 1, c), ctLem, excurly)));
    end;
    if Sender=FindDown then Hop(1) else if Sender=FindUp then Hop(-1);
    repeat
      if t=High(Heads) then begin
        if memotxt='' then memotxt:=TextToCompare(TextPages[l].Text, ctLem, excurly);
        txt:=Copy(memotxt, c, Length(fnd));
      end else
        if Length(Tables[t])=0 then txt:='' else if r>-1 then txt:=TextToCompare(Tables[t,l].HTMLs[r, c+Ord(ColType[t,c]=ctSolv)], ColType[t,c], excurly)
          else if LemFindBtn.Tag=1 then txt:='' else if CaseCheckBox.Checked then txt:=Tables[t,l].Id else txt:=LowerCase(Tables[t,l].Id);
      if Pos(fnd, txt)>0 then begin
        f:=True;
        if TabControl.TabIndex<>t then begin TabControl.TabIndex:=t;  TabControlChange(FindEdit);  end;
        if t<High(Heads) then begin
          if ListBox.ItemIndex<>l then begin ListBox.ItemIndex:=l;    ListBoxClick(nil);           end;
          if r>-1 then begin
            StringGrid.Row:=r+1;
            StringGrid.Col:=c;
          end else ft:=True;
        end else begin
          i:=1;
          if excurly then repeat
            if Pos(TextPages[l].Text[i], '{}')>0 then Inc(c);
            Inc(i);
          until i>c;
          UnitList.ItemIndex:=l;
          UnitListChange(nil);
          Memo.SelStart:=c-1;  Memo.SelLength:=Length(fnd);
          for i:=1 to Length(Memo.SelText) do if Pos(Memo.SelText[i], '{}')>0 then Memo.SelLength:=Memo.SelLength+1;
        end;
      end else Hop(2*Ord(Sender<>FindUp)-1);
    until f
      or ((TabControl.TabIndex<High(Heads)) and ((t=TabControl.TabIndex) and (l=ListBox.ItemIndex) and ((FoundTitle and (r=-1)) or (not FoundTitle and (r=StringGrid.Row-1))) and (c=StringGrid.Col)))
      or ((TabControl.TabIndex=High(Heads)) and ((t=TabControl.TabIndex) and (l=UnitList.ItemIndex) and (c=Memo.SelStart+Ord(Sender=FindUp))))
      or (Now>ti+2/86400);
  end;
  FindEdit.Color:=clSuccess[f or (Length(FindEdit.Text)=0), False];
  FoundTitle:=ft;
end; {FindEditChange}

procedure TTableMaker.LemFindBtnClick(Sender: TObject);
const capt: Array[0..2] of string = ('Englis&h', 'Lemiz&h', 'anyt&hing');
      color: Array[0..2] of TColor = ($d3eaf7, $e7f7d3, clWhite);
begin
  if Sender<>nil then LemFindBtn.Tag:=(LemFindBtn.Tag+1) mod 3;
  LemFindBtn.Caption:=capt[LemFindBtn.Tag];
  FindPanel2.Color:=color[LemFindBtn.Tag];
  SetLemFont(FindEdit.Font, LemFindBtn.Tag=1, 0);
  CaseCheckBox.Visible:=LemFindBtn.Tag<>1;
  AccentCheckBox.Visible:=LemFindBtn.Tag=1;
  VowelCheckBox.Visible:=LemFindBtn.Tag=1;
  ControlEnterExit(Sender);
  if Sender<>nil then FindEditChange(Sender);
end; {LemFindBtnClick}

procedure TTableMaker.SymbolClick(Sender: TObject);
begin
  Clipboard.Open;
  Clipboard.Clear;
  with TMenuItem(Sender) do
    if Symbols[tsyHint,Tag]='' then CopyRTFasUnicode(Symbols[tsyOut,Tag]) else CopyRTFasUnicode('<abbr title="'+Symbols[tsyHint,Tag]+'">'+Symbols[tsyOut,Tag]+'</abbr>');
  Clipboard.Close;
end; {SymbolClick}

procedure TTableMaker.CaseBtnClick(Sender: TObject);
var P: TPoint;
begin
  P:=LemBar.ClientToScreen(Point(CaseBtn.Left, CaseBtn.Top+CaseBtn.Height));
  CaseMenu.Popup(P.X, P.Y);
end; {CaseBtnClick}

procedure TTableMaker.CaseClick(Sender: TObject);
begin
  with TMenuItem(Sender) do CopyRTFasUnicode('<abbr class="gloss" title="'+CaseExtds[Tag]+'">'+CaseAbbrs[Tag]+'</abbr>');
end; {CaseClick}

{------------------------------------------------ Memo -----------------------------------------------------------}

procedure TTableMaker.FormMouseWheel(Sender: TObject; Shift: TShiftState; WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
begin
  if Shift=[ssCtrl] then
    if Notebook.PageIndex=1 then begin
      MemoFontSize:=Max(MemoFontSize+Sign(WheelDelta), -TableMaker.Font.Size);
      SetLemFont(Memo.Font, True, MemoFontSize);
    end else if ActiveControl=SolvMemo then SolvMemo.Font.Size:=Max(SolvMemo.Font.Size+Sign(WheelDelta), 1);
end; {FormMouseWheel}

procedure TTableMaker.UnitListChange(Sender: TObject);
var c, s: Integer;
begin
  with TextPages[UnitList.ItemIndex] do begin
    c:=Caret;                    s:=Selection;
    Memo.Lines.Text:=Text;       Memo.SelStart:=c;            Memo.SelLength:=s;
    TextTagsCombo.Text:=Header;  ColspanEdit.Value:=ColSpan;  RowspanEdit.Value:=RowSpan;
  end;
  if Visible and not FindEdit.Focused then Memo.SetFocus;
end; {UnitListChange}

procedure TTableMaker.CutMemoBtnClick(Sender: TObject);
begin
  CopyMemoBtnClick(CopyMemoBtn);
  DelMemoBtnClick(DelMemoBtn);
end; {CutMemoBtnClick}

function TTableMaker.SelCaretLine: Integer;
var l: Integer;
begin
  Result:=-1;  l:=-1;
  while l<Memo.SelStart do begin
    Inc(l, Length(Memo.Lines[Result+1])+2);  Inc(Result);
  end;
end; {SelCaretLine}

function TTableMaker.HTMLSpans: string;
begin
  if SpanPanel.Visible and (ColspanEdit.Value>1) then Result:=' colspan="'+ColspanEdit.Text+'"' else Result:='';
  if SpanPanel.Visible and (RowspanEdit.Value>1) then Result:=Result+' rowspan="'+RowspanEdit.Text+'"';
end; {HTMLSpans}

procedure TTableMaker.CopyMemoBtnClick(Sender: TObject);
var st, st2, tg, stout, errmsg: string;
    tm: TDateTime;
begin
  if CopyMemoBtn.Enabled then begin
    tm:=Now;
    Clipboard.Open;
    Clipboard.Clear;
    if Memo.SelLength=0 then st:=Memo.Lines[SelCaretLine] else st:=Memo.SelText;
    tg:=RemoveFrom(' ', TextTagsCombo.Text);
    if TBitBtn(Sender).Tag in [0, 2] then begin
      st2:=SubstituteStems(st, True, Options.Dictionary.Items);
      stout:='<'+StringReplace(TextTagsCombo.Text, '$', MakeLemId(st2), [rfReplaceAll])+' lang="x-lm"'+HTMLSpans+LemTranscript(st2, Symbols, errmsg)+'>'+IncludeBracketsEtc(st2, True)+'</'+tg+'>';
    end else stout:='';
    with TBitBtn(Sender) do if Tag in [1, 2] then stout:=stout+IfThen(Tag=2, IfThen(tg='p', #13#10#13#10, ' '))+'<'+TextTagsCombo.Text+HTMLSpans+' class="igloss">'+Glossing(st, errmsg)+'</'+tg+'>';
    CopyRTFasUnicode(HTML4or5(stout));
    Clipboard.Close;
    ErrorMessageBox(errmsg);
    ActiveControl:=Memo;
    MemoTimeLabel.Caption:='Copied in '+FloatToStrF((Now-tm)*86400, ffGeneral, 4, 2)+' s';
  end;
end; {CopyMemoBtnClick}

procedure TTableMaker.DelMemoBtnClick(Sender: TObject);
begin
  if Memo.SelLength=0 then Memo.Lines.Delete(SelCaretLine) else Memo.ClearSelection;
  ActiveControl:=Memo;
end; {DelMemoBtnClick}

procedure TTableMaker.RandomBtnClick(Sender: TObject);
begin
  Memo.Lines[SelCaretLine]:=RandomSentence;
end; {RandomBtnClick}

procedure TTableMaker.MemoChange(Sender: TObject);
begin
  TextPages[UnitList.ItemIndex].Text:=Memo.Lines.Text;
  MemoMouseUp(Sender, mbLeft, [], 0, 0);
end; {MemoChange}

procedure TTableMaker.MemoKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Shift=[ssCtrl] then case Key of
    65{A}: TMemo(Sender).SelectAll;
    48{0}, VK_Numpad0, 187{+}, VK_Add, 189{-}, VK_Subtract: if Sender=Memo then begin
      if Key in [48, VK_Numpad0] then MemoFontSize:=0 else Inc(MemoFontSize, 2*Ord(Key in [187, VK_Add])-1);
      SetLemFont(Memo.Font, True, MemoFontSize);
    end else if Key in [48, VK_Numpad0] then SolvMemo.Font.Size:=Font.Size else SolvMemo.Font.Size:=SolvMemo.Font.Size+2*Ord(Key in [187, VK_Add])-1;
  end else if (Sender=Memo) and (Shift=[ssAlt]) then case Key of
    67{C}: CopyMemoBtnClick(CopyMemoBtn);
    88{X}: CutMemoBtnClick(nil);
  end;
  if (Shift=[ssCtrl]) and (Key in [$31, $37..$39]) then StringGridKeyDown(Sender, Key, Shift);
end; {MemoKeyDown}

procedure TTableMaker.MemoKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  MemoMouseUp(Sender, mbLeft, [], 0, 0);
end;

procedure TTableMaker.MemoMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  CutMemoBtn.Enabled:=(Memo.SelLength>0) or (Memo.Lines[SelCaretLine]<>'');
  CopyMemoBtn.Enabled:=CutMemoBtn.Enabled;
  CopyMemoGlossBtn.Enabled:=CutMemoBtn.Enabled;
  CopyMemoBothBtn.Enabled:=CutMemoBtn.Enabled;
  DelMemoBtn.Enabled:=SelCaretLine<Memo.Lines.Count;
  with TextPages[UnitList.ItemIndex] do begin
    Caret:=Memo.SelStart;  Selection:=Memo.SelLength;
  end;
end; {MemoMouseUp}

procedure TTableMaker.TextTagsComboChange(Sender: TObject);
begin
  SpanPanel.Visible:=(TextTagsCombo.Text='td') or (TextTagsCombo.Text='th');
  with TextPages[UnitList.ItemIndex] do begin
    Header:=TextTagsCombo.Text;  ColSpan:=ColspanEdit.Value;   RowSpan:=RowspanEdit.Value;
  end;
  if Visible then CopyMemoBtnClick(CopyMemoBtn);
end; {TextTagsComboChange}

procedure TTableMaker.LemBoxClick(Sender: TObject);
begin
  StringGrid.Invalidate;
  SetLemFont(Memo.Font, True, MemoFontSize);
  SetLemFont(FindEdit.Font, LemFindBtn.Tag=1, 0);
end; {LemBoxClick}

{------------------------------------------------ Utilities -----------------------------------------------------------}

function CaseEnding(i: Integer): string;
const cv = 'aeyioöuü';
      cs: array[0..3] of string = ('', 'l', 'rh', 'r');
begin
  Result:=IfThen(i<32, cv[i mod 8 +1]+cs[i div 8])+IfThen(i=32, 'ng')+IfThen(i=33, 'm');
end; {CaseEnding}

function TTableMaker.ThisTable: TTable;
begin
  if TabControl.TabIndex=High(Heads) then Result:=nil else
    if Length(Tables[TabControl.TabIndex])>ListBox.ItemIndex then Result:=Tables[TabControl.TabIndex, ListBox.ItemIndex] else Result:=nil;
end; {ThisTable}

function TTableMaker.FullSolutionsPath: string;
begin
  Result:=SolutionsPath+'\'+IntToStr(UnitCombo.ItemIndex)+'_'+LowerCase(IdEdit.Text)+'_';
end; {FullSolutionsPath}

function TTableMaker.GetNoOfSolutions(A, B, Un: Integer; Id: string; out No, SolvCol: Integer): string;
var i: Integer;
begin
  No:=0;  SolvCol:=-1;
  Result:=SolutionsPath+'\'+IntToStr(Un)+'_'+LowerCase(Id)+'_';
  for i:=0 to ColCounts[TabControl.TabIndex]-1 do if ColType[TabControl.TabIndex, i]=ctSolv then SolvCol:=i;
  if SolvCol>-1 then for i:=A to B do if ThisTable.HTMLs[i-1, SolvCol]<>'' then Inc(No);
end; {GetNoOfSolutions}

procedure TTableMaker.SetLemFont(Font: TFont; CanLem: Boolean; SizeOffset: Integer);
begin
  if LemBox.Checked and CanLem then Font.Name:=LemFont else Font.Name:=SansSerif;
  if CanLem then Font.Size:=TableMaker.Font.Size+1+Ord(LemBox.Checked)+SizeOffset else Font.Size:=TableMaker.Font.Size;
end; {SetLemFont}

procedure TTableMaker.SetCaption;
begin
  if Notebook.PageIndex=1 then Caption:='Text - '+Application.Title else
    Caption:=Heads[TabControl.TabIndex]+IfThen(GridPanel.Visible, ': '+UnitCombo.Text
      +IfThen(IdEdit.Text<>'', '. '+IdEdit.Text))+' - '+Application.Title;
end; {SetCaption}

procedure TTableMaker.UpdateCasePopup;
var i, j: Integer;
    b: TMenuBreak;
begin
  for i:=CaseMenu.Items.Count-1 downto 0 do CaseMenu.Items.Delete(i);
  for j:=0 to Length(Symbols[0])-1 do if Symbols[tsyCode, j]='' then AddMenuItem(self, '-', 0, nil, CaseMenu.Items) else
    AddMenuItem(self, StringReplace(Symbols[tsyCode, j]+IfThen(Symbols[tsyCode, j]='', Symbols[tsyOut, j])+IfThen(Symbols[tsyHint, j]<>'', ' ('+Symbols[tsyHint, j]+')'), '&', '&&', [rfReplaceAll]),
      j, SymbolClick, CaseMenu.Items);
  if not CaseMenuBreak then AddMenuItem(self, '-', 0, nil, CaseMenu.Items);
  for i:=0 to High(CaseAbbrs) do begin
    b:=TMenuBreak(2*Ord(CaseMenuBreak and (i=0)));
    if CaseAbbrs[i]='' then AddMenuItem(self, '('+CaseEnding(i)+')', i, nil, CaseMenu.Items).Break:=b
      else AddMenuItem(self, UpperCase(CaseAbbrs[i])+' ('+CaseExtds[i]+', '+CaseEnding(i)+')', i, CaseClick, CaseMenu.Items).Break:=b;
    if i mod 8 = 7 then AddMenuItem(self, '-', 0, nil, CaseMenu.Items);
  end;
end; {UpdateCasePopup}

procedure TTableMaker.PopupPopup(Sender: TObject);
const EditSolText: array[False..True] of string = ('Edit', 'Create');
var p: TPoint;
    i: Integer;
begin
  with TWinControl(TPopupMenu(Sender).PopupComponent) do begin
    p:=ScreenToClient(Mouse.CursorPos);
    if (Sender=TablePopup) or not MultiSelectBtn.Down or not (StringGrid.MouseCoord(p.X, p.Y).Y in [StringGrid.Selection.Top..StringGrid.Selection.Bottom]) then
      SendMessage(Handle, WM_LButtonDown, 0, MakeLong(p.X, p.Y));
  end;
  if Sender=TablePopup then begin
    ListBoxClick(nil);
    for i:=1 to MarksHi do with TMenuItem(FindComponent('Mark'+IntToStr(i))) do begin
      Enabled:=ListBox.ItemIndex>-1;
      Checked:=Enabled and (ThisTable.Mark=i);
    end;
    Delete1.Enabled:=DelTableBtn.Enabled;
    Up1.Enabled:=UpTableBtn.Enabled;
    Down1.Enabled:=DownTableBtn.Enabled;
  end else begin {LinesPopup}
    StringGrid.MouseToCell(p.X, p.Y, p.X, p.Y);
    Sortsolutions1.Visible:=ColType[TabControl.TabIndex, p.X]=ctSolv;
    Editsolution1.Visible:=Sortsolutions1.Visible and (p.Y>0);
    if Editsolution1.Visible then EditSolution1.Caption:=EditSolText[ThisTable.HTMLs[p.Y-1, p.X+1]='']+' solution';
    Copytablefornutshell1.Enabled:=CopyNutshellBtn.Enabled;
    Suppressfornutshell1.Checked:=NutshellBtn.Down;
    SolvX:=p.x+1;  SolvY:=p.y-1;
    Pasteline1.Enabled:=Clipboard.HasFormat(CF_Rows[TabControl.TabIndex]);
    Up2.Enabled:=UpLineBtn.Enabled;
    Down2.Enabled:=DownLineBtn.Enabled;
  end;
end; {PopupPopup}

procedure TTableMaker.UpdateTextPages;
var i, u: Integer;
begin
  u:=UnitList.ItemIndex;
  UnitList.Items.Assign(UnitCombo.Items);
  UnitList.Items.Insert(0, '(sandbox)');
  with UnitList.Items do for i:=Count to Length(TextPages)-1 do TextPages[Count-1].Text:=TextPages[Count-1].Text+#10#13#10#13+TextPages[i].Text;
  SetLength(TextPages, UnitList.Items.Count);
  UnitList.ItemIndex:=Min(u, UnitList.Items.Count-1);
end; {UpdateTextPages}

function TTableMaker.HTML4or5(S: string): string;
begin
  if not HTML5Image.Visible then S:=StringReplace(StringReplace(S, '<mark>', '<em>', [rfReplaceAll]), '</mark>', '</em>', [rfReplaceAll]);
  Result:=S;
end; {HTML4or5}

procedure TTableMaker.ErrorMessageBox(Err: string);
begin
  if Err<>'' then TangoMessageBox(Err, TMsgDlgType(IfThen(Pos('error', Err)>0, Integer(mtError), Integer(mtInformation))), [mbOK], '');
end; {ErrorMessageBox}

(*
* (eat)y\[(identity)y\](ask)à (sweet)Ì{y}. doesn’t bracket (post)stem
* Glossing (function TTableMaker.Glossing)
  + Highlight case EXCLUDING part/qual
    - (speak)à{\(m\)}. breaks glossing, also \[ \] \{ \}, also primary suffix
  + Highlighting of complete PIIn doesn’t always work: (wine)ìl (drink)ýy zeí{ ýmym zìi}. OK, (wine)ìl (drink)ýy zeí {ýmym zìi}. NOT OK. (ev. dummy stem, e.g. "@ýmym"?)
  + Emdashes
*)

end.

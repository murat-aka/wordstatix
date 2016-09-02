// ***********************************************************************
// ***********************************************************************
// WordStatix 1.7
// Author and copyright: Massimo Nardello, Modena (Italy) 2016.
// Free software released under GPL licence version 3 or later.
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version. You can read the version 3
// of the Licence in http://www.gnu.org/licenses/gpl-3.0.txt
// or in the file Licence.txt included in the files of the
// source code of this software.
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.
// ***********************************************************************
// ***********************************************************************
unit unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, Grids,
  StdCtrls, ExtCtrls, ComCtrls, Menus, LazUTF8, TAGraph, TASources, TASeries,
  TATools, IniFiles, Clipbrd, zipper, LCLIntf, CheckLst, Types, LazFileUtils;

type

  TRecWordList = record
    stRecWord: string;
    iRecFreq: integer;
    stRecWordPos: string;
    stRecBookPos: string;
  end;

  TRecWordTextual = record
    stRecWordTextual: string;
  end;

  { TfmMain }

  TfmMain = class(TForm)
    apAppProp: TApplicationProperties;
    bnDeselConc: TButton;
    bnFindFirst: TButton;
    bnFindNext: TButton;
    bnReplace: TButton;
    bnReplaceAll: TButton;
    bnDeselDiag: TButton;
    cbCollatedSort: TCheckBox;
    cbComboDiag1: TComboBox;
    cbComboDiag2: TComboBox;
    cbComboDiag3: TComboBox;
    cbComboDiag4: TComboBox;
    cbComboDiag5: TComboBox;
    chChartBarSeries4: TBarSeries;
    chChartBarSeries5: TBarSeries;
    clDiagBookmark: TCheckListBox;
    csChartSource4: TListChartSource;
    csChartSource5: TListChartSource;
    ctChartToolset: TChartToolset;
    chChart: TChart;
    chChartBarSeries1: TBarSeries;
    chChartBarSeries2: TBarSeries;
    chChartBarSeries3: TBarSeries;
    cbSkipNumbers: TCheckBox;
    csChartSource2: TListChartSource;
    csChartSource3: TListChartSource;
    edFltStart: TEdit;
    edFltEnd: TEdit;
    edFind: TEdit;
    edLocate: TEdit;
    edReplace: TEdit;
    edWordsCont: TEdit;
    lbDiagBookmark: TLabel;
    lbDiag3: TLabel;
    lbDiag2: TLabel;
    lbDiag1: TLabel;
    lbBookmarks: TListBox;
    lbDiag4: TLabel;
    lbDiag5: TLabel;
    lbLocate: TLabel;
    lbReplace: TLabel;
    lbFind: TLabel;
    lbFltStart: TLabel;
    lbFltEnd: TLabel;
    lbWordsCont: TLabel;
    lbContext: TLabel;
    lbNoWord: TLabel;
    csChartSource1: TListChartSource;
    lsContext: TListBox;
    miLine1: TMenuItem;
    miFileSaveConc: TMenuItem;
    miFileOpenConc: TMenuItem;
    miStatisticSortFreq: TMenuItem;
    miStatisticSortName: TMenuItem;
    miLine9: TMenuItem;
    miLine5: TMenuItem;
    miConcordanceJoin: TMenuItem;
    miDiagramShowGrid: TMenuItem;
    miDiagramAllWordNoBook: TMenuItem;
    miConcordanceRefreshGrid: TMenuItem;
    miLine6: TMenuItem;
    miDiagramAllWordsBook: TMenuItem;
    miLine13: TMenuItem;
    miDiagramShowVal: TMenuItem;
    miDiaZoomIn: TMenuItem;
    miDiaZoomOut: TMenuItem;
    miDiaZoomNormal: TMenuItem;
    miLine14: TMenuItem;
    miLine12: TMenuItem;
    miDiagramSave: TMenuItem;
    miDiagramSingleWordsBook: TMenuItem;
    miDiagram: TMenuItem;
    miConcordanceRemove: TMenuItem;
    miManual: TMenuItem;
    miLine15: TMenuItem;
    miFileNew: TMenuItem;
    miStatisticSave: TMenuItem;
    miStatisticCreate: TMenuItem;
    miLine10: TMenuItem;
    miStatistic: TMenuItem;
    miFileSaveAs: TMenuItem;
    miLine8: TMenuItem;
    miConcordanceDelCont: TMenuItem;
    miLine2: TMenuItem;
    miFileUpdBookmark: TMenuItem;
    miFileSetBookmark: TMenuItem;
    miConcordanceOpenSkip: TMenuItem;
    miConcordanceSaveSkip: TMenuItem;
    miLine4: TMenuItem;
    miConcodanceShowSelected: TMenuItem;
    miFileSave: TMenuItem;
    miLine7: TMenuItem;
    miConcordanceSaveRep: TMenuItem;
    miCopyrightForm: TMenuItem;
    miCopyright: TMenuItem;
    miLine3: TMenuItem;
    miConcordanceAddSkip: TMenuItem;
    miConcordanceCreate: TMenuItem;
    miFileOpen: TMenuItem;
    miLine0: TMenuItem;
    miFileExit: TMenuItem;
    mmMenu: TMainMenu;
    miConcordance: TMenuItem;
    miFile: TMenuItem;
    meText: TMemo;
    meSkipList: TMemo;
    odOpenDialog: TOpenDialog;
    pnDiagram: TPanel;
    pnListBookmark: TPanel;
    pnTextBottom: TPanel;
    pcMain: TPageControl;
    rgCondFlt: TRadioGroup;
    rgSortBy: TRadioGroup;
    sbDiagram: TScrollBox;
    sdSaveDialog: TSaveDialog;
    sgWordList: TStringGrid;
    sbStatusBar: TStatusBar;
    sgStatistic: TStringGrid;
    spDiagram: TSplitter;
    spText: TSplitter;
    tsDiagram: TTabSheet;
    tsStatistic: TTabSheet;
    tsFile: TTabSheet;
    tsConcordance: TTabSheet;
    udWordsCont: TUpDown;
    procedure apAppPropException(Sender: TObject; E: Exception);
    procedure bnDeselConcClick(Sender: TObject);
    procedure bnFindFirstClick(Sender: TObject);
    procedure bnFindNextClick(Sender: TObject);
    procedure bnReplaceAllClick(Sender: TObject);
    procedure bnReplaceClick(Sender: TObject);
    procedure bnDeselDiagClick(Sender: TObject);
    procedure cbCollatedSortChange(Sender: TObject);
    procedure cbComboDiag2KeyDown(Sender: TObject; var Key: word;
      Shift: TShiftState);
    procedure cbComboDiag3KeyDown(Sender: TObject; var Key: word;
      Shift: TShiftState);
    procedure cbComboDiag4KeyDown(Sender: TObject; var Key: word;
      Shift: TShiftState);
    procedure cbComboDiag5KeyDown(Sender: TObject; var Key: word;
      Shift: TShiftState);
    procedure clDiagBookmarkKeyUp(Sender: TObject; var Key: word;
      Shift: TShiftState);
    procedure edFltEndExit(Sender: TObject);
    procedure edFltStartExit(Sender: TObject);
    procedure edLocateKeyPress(Sender: TObject; var Key: char);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure FormShow(Sender: TObject);
    procedure lbBookmarksMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: integer);
    procedure meSkipListExit(Sender: TObject);
    procedure meTextChange(Sender: TObject);
    procedure meTextKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure meTextMouseWheel(Sender: TObject; Shift: TShiftState;
      WheelDelta: integer; MousePos: TPoint; var Handled: boolean);
    procedure miConcodanceShowSelectedClick(Sender: TObject);
    procedure miConcordanceAddSkipClick(Sender: TObject);
    procedure miConcordanceCreateClick(Sender: TObject);
    procedure miConcordanceJoinClick(Sender: TObject);
    procedure miFileOpenConcClick(Sender: TObject);
    procedure miConcordanceOpenSkipClick(Sender: TObject);
    procedure miConcordanceRefreshGridClick(Sender: TObject);
    procedure miConcordanceRemoveClick(Sender: TObject);
    procedure miFileSaveConcClick(Sender: TObject);
    procedure miConcordanceSaveRepClick(Sender: TObject);
    procedure miConcordanceDelContClick(Sender: TObject);
    procedure miConcordanceSaveSkipClick(Sender: TObject);
    procedure miCopyrightFormClick(Sender: TObject);
    procedure miDiagramAllWordNoBookClick(Sender: TObject);
    procedure miDiagramShowGridClick(Sender: TObject);
    procedure miDiagramSingleWordsBookClick(Sender: TObject);
    procedure miDiagramAllWordsBookClick(Sender: TObject);
    procedure miDiagramSaveClick(Sender: TObject);
    procedure miDiagramShowValClick(Sender: TObject);
    procedure miDiaZoomInClick(Sender: TObject);
    procedure miDiaZoomNormalClick(Sender: TObject);
    procedure miDiaZoomOutClick(Sender: TObject);
    procedure miFileNewClick(Sender: TObject);
    procedure miManualClick(Sender: TObject);
    procedure miStatisticCreateClick(Sender: TObject);
    procedure miFileSaveAsClick(Sender: TObject);
    procedure miFileSetBookmarkClick(Sender: TObject);
    procedure miFileUpdBookmarkClick(Sender: TObject);
    procedure miFileExitClick(Sender: TObject);
    procedure miFileOpenClick(Sender: TObject);
    procedure miFileSaveClick(Sender: TObject);
    procedure miStatisticSaveClick(Sender: TObject);
    procedure miStatisticSortFreqClick(Sender: TObject);
    procedure miStatisticSortNameClick(Sender: TObject);
    procedure pcMainChange(Sender: TObject);
    procedure pcMainChanging(Sender: TObject; var AllowChange: boolean);
    procedure rgSortBySelectionChanged(Sender: TObject);
    procedure sgStatisticPrepareCanvas(Sender: TObject; aCol, aRow: integer;
      aState: TGridDrawState);
    procedure sgStatisticSelectCell(Sender: TObject; aCol, aRow: integer;
      var CanSelect: boolean);
    procedure sgWordListKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure sgWordListSelection(Sender: TObject; aCol, aRow: integer);
    procedure tsDiagramResize(Sender: TObject);
    procedure udWordsContChanging(Sender: TObject; var AllowChange: boolean);
  private
    procedure CompileGrid;
    procedure CreateConc(stStartText: string);
    function CreateContext(GridRow: integer; blSetCursor: boolean): string;
    procedure AddSkipWord;
    function CheckFilter(stWord: string): boolean;
    function CleanField(myField: string): string;
    function CleanXML(stXMLText: string): string;
    procedure CreateDiagramAllWordsBook;
    procedure CreateDiagramAllWordsNoBook;
    procedure CreateDiagramSingleWordsBook;
    procedure CreateStatistic;
    procedure DisableMenuItems;
    procedure EnableMenuItems;
    function FindNextWord(blMessage: boolean): boolean;
    procedure OpenConcordance;
    procedure SaveConcordance;
    procedure SaveReport(stFileName: string);
    procedure CreateBookmarks(blSetCursor: boolean);
    procedure SortWordFreq(var arWordList: array of TRecWordList;
      flField: shortint);
    { private declarations }
  public
    { public declarations }
  end;

var
  fmMain: TfmMain;
  arWordList: array of TRecWordList;
  arWordTextual: array of TRecWordTextual;
  myHomeDir: string;
  myConfigFile: string;
  stFileName: string;
  iWordsUsed: integer = 0;
  iWordsTotal: integer = 0;
  iWordsStartTot: integer = 0;
  ttStart, ttEnd: TTime;
  blGridConcMod: boolean = False;
  blStopConcordance: boolean = False;
  blConInProc: boolean = False;
  blTextModSave: boolean = False;
  blTextModConc: boolean = False;
  stDiagTitle: string;

implementation

{$R *.lfm}

uses Copyright;

{ TfmMain }

// ********************************************************** //
// ***************** Events of the main form **************** //
// ********************************************************** //

procedure TfmMain.FormCreate(Sender: TObject);
var
  MyIni: TIniFile;
begin
  // Set data on start
  fmMain.Color := fmMain.Color;
  if fmMain.Color <> clDefault then
  begin
    fmMain.sgWordList.FixedColor := fmMain.Color;
    fmMain.sgStatistic.FixedColor := fmMain.Color;
    lsContext.Color := fmMain.Color;
  end;
  sgWordList.FocusRectVisible := False;
  sgStatistic.FocusRectVisible := False;
  lbBookmarks.ScrollWidth := 0;
  lsContext.ScrollWidth := 0;
  clDiagBookmark.ScrollWidth := 0;
  miFileSave.Enabled := False;
  {$ifdef Linux}
  myHomeDir := GetEnvironmentVariable('HOME') + DirectorySeparator +
    '.config' + DirectorySeparator + 'wordstatix' + DirectorySeparator;
  myConfigFile := 'wordstatix';
  {$endif}
  {$ifdef Win32}
  myHomeDir := GetAppConfigDir(False);
  myConfigFile := 'wordstatix.ini';
  fmMain.Color := clWhite;
  lbBookmarks.Color := clWhite;
  clDiagBookmark.Color := clWhite;
  {$endif}
  {$ifdef Darwin}
  myHomeDir := GetEnvironmentVariable('HOME') + DirectorySeparator +
    'Library' + DirectorySeparator + 'Preferences' + DirectorySeparator;
  myConfigFile := 'wordstatix.plist';
  fmMain.Color := clWhite;
  {$endif}
  if DirectoryExists(myHomeDir) = False then
  begin
    CreateDirUTF8(myHomeDir);
  end;
  if FileExistsUTF8(myHomeDir + 'skipwords') then
  begin
    meSkipList.Lines.LoadFromFile(myHomeDir + 'skipwords');
  end;
  if FileExistsUTF8(myHomeDir + myConfigFile) then
    try
      MyIni := TIniFile.Create(myHomeDir + myConfigFile);
      if MyIni.ReadString('wordstatix', 'maximize', '') = 'true' then
      begin
        fmMain.WindowState := wsMaximized;
      end
      else
      begin
        fmMain.Top := MyIni.ReadInteger('wordstatix', 'top', 0);
        fmMain.Left := MyIni.ReadInteger('wordstatix', 'left', 0);
        if MyIni.ReadInteger('wordstatix', 'width', 0) > 100 then
          fmMain.Width := MyIni.ReadInteger('wordstatix', 'width', 0)
        else
          fmMain.Width := 1000;
        if MyIni.ReadInteger('wordstatix', 'heigth', 0) > 100 then
          fmMain.Height := MyIni.ReadInteger('wordstatix', 'heigth', 0)
        else
          fmMain.Height := 600;
      end;
      if MyIni.ReadInteger('wordstatix', 'sortby', 0) > -1 then
      begin
        rgSortBy.ItemIndex := MyIni.ReadInteger('wordstatix', 'sortby', 0);
      end;
      if MyIni.ReadString('wordstatix', 'wordscont', '') <> '' then
      begin
        edWordsCont.Text := MyIni.ReadString('wordstatix', 'wordscont', '6');
      end;
      if MyIni.ReadString('wordstatix', 'collatesort', '') = 't' then
      begin
        cbCollatedSort.Checked := True;
      end;
      if MyIni.ReadString('wordstatix', 'skipnum', '') = 't' then
      begin
        cbSkipNumbers.Checked := True;
      end;
    finally
      MyIni.Free;
    end;
end;

procedure TfmMain.FormClose(Sender: TObject; var CloseAction: TCloseAction);
var
  MyIni: TIniFile;
begin
  // Save data on close
  blStopConcordance := True;
  meSkipList.Lines.SaveToFile(myHomeDir + 'skipwords');
  try
    MyIni := TIniFile.Create(myHomeDir + myConfigFile);
    if fmMain.WindowState = wsMaximized then
    begin
      MyIni.WriteString('wordstatix', 'maximize', 'true');
    end
    else
    begin
      MyIni.WriteString('wordstatix', 'maximize', 'false');
      MyIni.WriteInteger('wordstatix', 'top', fmMain.Top);
      MyIni.WriteInteger('wordstatix', 'left', fmMain.Left);
      MyIni.WriteInteger('wordstatix', 'width', fmMain.Width);
      MyIni.WriteInteger('wordstatix', 'heigth', fmMain.Height);
    end;
    MyIni.WriteString('wordstatix', 'wordscont', edWordsCont.Text);
    MyIni.WriteInteger('wordstatix', 'sortby', rgSortBy.ItemIndex);
    if cbCollatedSort.Checked = True then
    begin
      MyIni.WriteString('wordstatix', 'collatesort', 't');
    end
    else
    begin
      MyIni.WriteString('wordstatix', 'collatesort', 'f');
    end;
    if cbSkipNumbers.Checked = True then
    begin
      MyIni.WriteString('wordstatix', 'skipnum', 't');
    end
    else
    begin
      MyIni.WriteString('wordstatix', 'skipnum', 'f');
    end;
  finally
    MyIni.Free;
  end;
end;

procedure TfmMain.FormKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  // Select tab
  if ((Key = Ord('1')) and (Shift = [ssCtrl])) then
  begin
    pcMain.ActivePage := tsFile;
    key := 0;
  end
  else if ((Key = Ord('2')) and (Shift = [ssCtrl])) then
  begin
    pcMain.ActivePage := tsConcordance;
    key := 0;
  end
  else if ((Key = Ord('3')) and (Shift = [ssCtrl])) then
  begin
    pcMain.ActivePage := tsStatistic;
    key := 0;
  end
  else if ((Key = Ord('4')) and (Shift = [ssCtrl])) then
  begin
    pcMain.ActivePage := tsDiagram;
    key := 0;
  end;
  // Move the diagram
  if ((pcMain.ActivePage = tsDiagram) and (chChart.Visible = True)) then
  begin
    if key = 37 then
    begin
      sbDiagram.HorzScrollBar.Position :=
        sbDiagram.HorzScrollBar.Position - 50;
      key := 0;
    end
    else if key = 39 then
    begin
      sbDiagram.HorzScrollBar.Position :=
        sbDiagram.HorzScrollBar.Position + 50;
      key := 0;
    end;
  end;
  // Stop concordance
  if ((Key = Ord('H')) and (Shift = [ssShift, ssCtrl])) then
  begin
    blStopConcordance := True;
  end;
end;

procedure TfmMain.FormShow(Sender: TObject);
begin
  // Focus on text
  pcMain.ActivePage := tsFile;
  meText.SetFocus;
end;

procedure TfmMain.apAppPropException(Sender: TObject; E: Exception);
begin
  // Error handling
  MessageDlg('Operation not correct.', mtError, [mbOK], 0);
end;

procedure TfmMain.lbBookmarksMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: integer);
begin
  // Go to bookmark
  if lbBookmarks.Items.Count > 0 then
  begin
    meText.SelStart := UTF8Pos('[[' + lbBookmarks.Items[lbBookmarks.ItemIndex] +
      ']]', meText.Text) + 1;
    meText.SelLength := UTF8Length(lbBookmarks.Items[lbBookmarks.ItemIndex]);
    meText.SetFocus;
  end;
end;

procedure TfmMain.pcMainChange(Sender: TObject);
begin
  // Selecting the tsConcordance the sdGrid do not accept focus, so...
  if pcMain.ActivePage = tsConcordance then
    tsConcordance.SetFocus;
end;

procedure TfmMain.pcMainChanging(Sender: TObject; var AllowChange: boolean);
begin
  // Prevent change tab during concordance
  if blConInProc = True then
  begin
    AllowChange := False;
  end
  else
  begin
    AllowChange := True;
  end;
end;

procedure TfmMain.rgSortBySelectionChanged(Sender: TObject);
begin
  // Changed settings for concordance
  // also for rgCondFlt
  blTextModConc := True;
end;

procedure TfmMain.cbCollatedSortChange(Sender: TObject);
begin
  // Changed settings for concordance
  // also for all other Edit and Memo components
  blTextModConc := True;
end;

procedure TfmMain.sgWordListKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  // Select and deselect the word
  if key = 32 then
  begin
    sgWordList.Col := 0;
    if sgWordList.Cells[2, sgWordList.Row] = '0' then
    begin
      sgWordList.Cells[2, sgWordList.Row] := '1';
    end
    else
    begin
      sgWordList.Cells[2, sgWordList.Row] := '0';
    end;
    if sgWordList.Row < sgWordList.RowCount - 1 then
    begin
      sgWordList.Row := sgWordList.Row + 1;
    end;
  end;
end;

procedure TfmMain.bnDeselConcClick(Sender: TObject);
var
  i: integer;
begin
  // Select and deselect all words
  if sgWordList.RowCount > 1 then
  begin
    if sgWordList.Cells[2, 1] = '1' then
    begin
      for i := 1 to sgWordList.RowCount - 1 do
      begin
        sgWordList.Cells[2, i] := '0';
      end;
    end
    else
    begin
      for i := 1 to sgWordList.RowCount - 1 do
      begin
        sgWordList.Cells[2, i] := '1';
      end;
    end;
  end;
end;

procedure TfmMain.sgWordListSelection(Sender: TObject; aCol, aRow: integer);
begin
  // Create the list of context
  if sgWordList.RowCount > 1 then
  begin
    if sgWordList.RowHeights[sgWordList.Row] > 0 then
    begin
      lsContext.Clear;
      lsContext.Items.Text := CreateContext(sgWordList.Row, True);
      {$ifdef Win32}
      // Due to a bug in Windows...
      if lsContext.Items[0] = '' then
        lsContext.Items.Delete(0);
      {$endif}
      lsContext.ItemIndex := 0;
    end;
  end;
end;

procedure TfmMain.sgStatisticPrepareCanvas(Sender: TObject;
  aCol, aRow: integer; aState: TGridDrawState);
begin
  // Font of the last row and second col in Statistic
  if aRow = sgStatistic.RowCount - 1 then
  begin
    sgStatistic.Canvas.Font.Style := [fsBold];
    sgStatistic.Canvas.Brush.Color := clBtnFace;
  end
  else
  if aCol = 1 then
  begin
    sgStatistic.Canvas.Font.Style := [fsBold];
  end;
end;

procedure TfmMain.sgStatisticSelectCell(Sender: TObject; aCol, aRow: integer;
  var CanSelect: boolean);
begin
  // Avoid select last row
  if aRow = sgStatistic.RowCount - 1 then
  begin
    CanSelect := False;
  end;
end;

procedure TfmMain.udWordsContChanging(Sender: TObject; var AllowChange: boolean);
begin
  // Create the list of context
  if sgWordList.RowCount > 1 then
  begin
    if sgWordList.RowHeights[sgWordList.Row] > 0 then
    begin
      lsContext.Clear;
      lsContext.Items.Text := CreateContext(sgWordList.Row, True);
      {$ifdef Win32}
      // Due to a bug in Windows...
      if lsContext.Items[0] = '' then
        lsContext.Items.Delete(0);
      {$endif}
      lsContext.ItemIndex := 0;
    end;
  end;
end;

procedure TfmMain.meSkipListExit(Sender: TObject);
begin
  // Clean skip words on exit
  meSkipList.Text := CleanField(meSkipList.Text);
end;

procedure TfmMain.meTextChange(Sender: TObject);
begin
  // Set the text as modified
  blTextModSave := True;
  miFileSave.Enabled := blTextModSave;
  blTextModConc := True;
end;

procedure TfmMain.edFltStartExit(Sender: TObject);
begin
  // Clean filter start with on exit
  edFltStart.Text := CleanField(edFltStart.Text);
end;

procedure TfmMain.edLocateKeyPress(Sender: TObject; var Key: char);
var
  i: integer;
begin
  // Locate word in concordance
  if key = #13 then
  begin
    if edLocate.Text <> '' then
    begin
      for i := 1 to sgWordList.RowCount - 1 do
      begin
        if UTF8LowerCase(UTF8Copy(sgWordList.Cells[0, i], 1,
          UTF8Length(edLocate.Text))) = UTF8LowerCase(edLocate.Text) then
        begin
          if sgWordList.RowHeights[i] > 0 then
          begin
            sgWordList.Row := i;
          end;
          Break;
        end;
      end;
    end;
  end;
end;

procedure TfmMain.edFltEndExit(Sender: TObject);
begin
  // Clean filter end on exit
  edFltEnd.Text := CleanField(edFltEnd.Text);
end;

procedure TfmMain.meTextMouseWheel(Sender: TObject; Shift: TShiftState;
  WheelDelta: integer; MousePos: TPoint; var Handled: boolean);
begin
  // Change Zoom with mouse weel
  if Shift = [ssCtrl] then
  begin
    if WheelDelta > 0 then
    begin
      if meText.Font.Size < 48 then
        meText.Font.Size := meText.Font.Size + 1;
    end
    else
    begin
      if meText.Font.Size > 6 then
        meText.Font.Size := meText.Font.Size - 1;
    end;
  end;
end;

procedure TfmMain.meTextKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  // Change Zoom with keys
  if ((Key = 187) and (Shift = [ssShift, ssCtrl])) then
  begin
    if meText.Font.Size < 48 then
      meText.Font.Size := meText.Font.Size + 1;
  end
  else if ((Key = 189) and (Shift = [ssShift, ssCtrl])) then
  begin
    if meText.Font.Size > 6 then
      meText.Font.Size := meText.Font.Size - 1;
  end
  else if ((Key = Ord('0')) and (Shift = [ssCtrl])) then
  begin
    meText.Font.Size := 14;
  end
  else if ((Key = Ord('P')) and (Shift = [ssCtrl, ssShift])) then
  begin
    if meText.Text = '' then
    begin
      Abort;
    end;
    if MessageDlg('Compact all paragraphs not separated by an empty line, ' +
      'clean the text from double spaces and check the spaces ' +
      'after the punctuation marks?', mtConfirmation, [mbOK, mbCancel], 0) =
      mrCancel then
      Abort;
    try
      Screen.Cursor := crHourGlass;
      Application.ProcessMessages;
      meText.Text := StringReplace(meText.Text, '[[', #2, [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, '[', ' [', [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, '(', ' (', [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, #2, '[[', [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, '[[', ' [[', [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, '.', '. ', [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, ',', ', ', [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, ';', '; ', [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, ':', ': ', [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, '!', '! ', [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, '?', '? ', [rfReplaceAll]);
      while UTF8Pos(LineEnding + LineEnding + LineEnding, meText.Text) > 0 do
      begin
        meText.Text := StringReplace(meText.Text, LineEnding +
          LineEnding + LineEnding, LineEnding + LineEnding, [rfReplaceAll]);
      end;
      meText.Text := StringReplace(meText.Text, LineEnding + LineEnding,
        #2, [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, LineEnding, ' ', [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, #2, LineEnding +
        LineEnding, [rfReplaceAll]);
      while UTF8Pos('  ', meText.Text) > 0 do
      begin
        meText.Text := StringReplace(meText.Text, '  ', ' ', [rfReplaceAll]);
      end;
      meText.Text := StringReplace(meText.Text, LineEnding + ' [[',
        LineEnding + '[[', [rfReplaceAll]);
    finally
      Screen.Cursor := crDefault;
    end;
  end;
end;

procedure TfmMain.bnFindFirstClick(Sender: TObject);
begin
  // Find first
  if edFind.Text <> '' then
  begin
    if UTF8Pos(UTF8LowerCase(edFind.Text), UTF8LowerCase(meText.Text), 1) > 0 then
    begin
      meText.SelStart := UTF8Pos(UTF8LowerCase(edFind.Text),
        UTF8LowerCase(meText.Text), 1) - 1;
      meText.SelLength := UTF8Length(edFind.Text);
      meText.SetFocus;
    end
    else
    begin
      MessageDlg('No recurrences of the text to look for.',
        mtInformation, [mbOK], 0);
    end;
  end;
end;

procedure TfmMain.bnFindNextClick(Sender: TObject);
begin
  // Find next
  FindNextWord(True);
end;

procedure TfmMain.bnReplaceClick(Sender: TObject);
var
  stClip: string;
begin
  // Replace selection
  // Use clipboard because other methods
  // produce an unwanted vertical scrolling
  if meText.SelLength = 0 then
  begin
    MessageDlg('No text is selected.', mtWarning, [mbOK], 0);
    Abort;
  end;
  if edReplace.Text <> '' then
  begin
    stClip := Clipboard.AsText;
    Clipboard.AsText := edReplace.Text;
    meText.PasteFromClipboard;
    Clipboard.AsText := stClip;
    if edFind.Text <> '' then
    begin
      bnFindNextClick(nil);
    end;
  end;
end;

procedure TfmMain.bnReplaceAllClick(Sender: TObject);
var
  stText: string;
  i, iCount: integer;
begin
  // Replace all
  if ((edFind.Text <> '') and (edReplace.Text <> '')) then
  begin
    if MessageDlg('Replace all the recurrences of "' + edFind.Text +
      '" with "' + edReplace.Text + '" in the current text?',
      mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
    begin
      Abort;
    end;
    Screen.Cursor := crHourGlass;
    Application.ProcessMessages;
    stText := meText.Text;
    i := 1;
    iCount := 0;
    try
      while UTF8Pos(UTF8LowerCase(edFind.Text), UTF8LowerCase(stText), i) > 0 do
      begin
        i := UTF8Pos(UTF8LowerCase(edFind.Text), UTF8LowerCase(stText), i);
        stText := UTF8Copy(stText, 1, i - 1) + edReplace.Text +
          UTF8Copy(stText, i + UTF8Length(edFind.Text), UTF8Length(stText));
        i := i + UTF8Length(edReplace.Text);
        Inc(iCount);
        Application.ProcessMessages;
      end;
      meText.Text := stText;
    finally
      Screen.Cursor := crDefault;
    end;
    MessageDlg('The word looked for has been replaced ' + IntToStr(iCount) +
      ' times.', mtInformation, [mbOK], 0);
  end;
end;

procedure TfmMain.tsDiagramResize(Sender: TObject);
begin
  // Resize the diagram
  chChart.Width := sbDiagram.Width - 3;
  chChart.Height := sbDiagram.Height - 10;
end;

procedure TfmMain.bnDeselDiagClick(Sender: TObject);
begin
  // Select and deselect all bookmarks
  if clDiagBookmark.Count > 0 then
  begin
    if clDiagBookmark.Checked[0] = False then
    begin
      clDiagBookmark.CheckAll(cbChecked, False, False);
    end
    else
    begin
      clDiagBookmark.CheckAll(cbUnchecked, False, False);
    end;
  end;
end;

procedure TfmMain.cbComboDiag2KeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  // Clear the word field
  if ((Key = 46) or (key = 8)) then
    cbComboDiag2.Text := '';
end;

procedure TfmMain.cbComboDiag3KeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  // Clear the word field
  if ((Key = 46) or (key = 8)) then
    cbComboDiag3.Text := '';
end;

procedure TfmMain.cbComboDiag4KeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  // Clear the word field
  if ((Key = 46) or (key = 8)) then
    cbComboDiag4.Text := '';
end;

procedure TfmMain.cbComboDiag5KeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  // Clear the word field
  if ((Key = 46) or (key = 8)) then
    cbComboDiag5.Text := '';
end;

procedure TfmMain.clDiagBookmarkKeyUp(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  // Move on after space bar
  if key = 32 then
  begin
    if clDiagBookmark.ItemIndex < clDiagBookmark.Items.Count - 1 then
    begin
      clDiagBookmark.ItemIndex := clDiagBookmark.ItemIndex + 1;
    end;
  end;
end;

// ********************************************************** //
// ****************** Events of menu items ****************** //
// ********************************************************** //

procedure TfmMain.miFileNewClick(Sender: TObject);
begin
  // New
  if meText.Text <> '' then
  begin
    if MessageDlg('Create a new text and a new concordance ' +
      'closing the existing ones?', mtConfirmation, [mbOK, mbCancel], 0) =
      mrCancel then
    begin
      Abort;
    end;
  end;
  SetLength(arWordList, 0);
  SetLength(arWordTextual, 0);
  meText.Clear;
  lbBookmarks.Clear;
  edFind.Clear;
  edReplace.Clear;
  sgWordList.RowCount := 1;
  lsContext.Clear;
  edFltStart.Clear;
  edFltEnd.Clear;
  sgStatistic.RowCount := 0;
  sgStatistic.ColCount := 0;
  chChart.Visible := False;
  chChartBarSeries1.Active := False;
  chChartBarSeries2.Active := False;
  chChartBarSeries3.Active := False;
  chChartBarSeries4.Active := False;
  chChartBarSeries5.Active := False;
  cbComboDiag1.Clear;
  cbComboDiag2.Clear;
  cbComboDiag3.Clear;
  cbComboDiag4.Clear;
  cbComboDiag5.Clear;
  clDiagBookmark.Clear;
  pcMain.ActivePage := tsFile;
  sbStatusBar.SimpleText := 'No active concordance.';
end;

procedure TfmMain.miFileOpenClick(Sender: TObject);
var
  myZip: TUnZipper;
  myList, stFileOrig: TStringList;
  i: integer;
begin
  // Open file
  pcMain.ActivePage := tsFile;
  if ((meText.Text <> '') and (blTextModSave = True)) then
  begin
    if MessageDlg('The current text has been changed but has not been saved. ' +
      'Reject the changes and open a new text?', mtConfirmation,
      [mbOK, mbCancel], 0) = mrCancel then
    begin
      Abort;
    end;
  end;
  SetLength(arWordList, 0);
  SetLength(arWordTextual, 0);
  meText.Clear;
  lbBookmarks.Clear;
  edFind.Clear;
  edReplace.Clear;
  sgWordList.RowCount := 1;
  lsContext.Clear;
  edFltStart.Clear;
  edFltEnd.Clear;
  sgStatistic.RowCount := 0;
  sgStatistic.ColCount := 0;
  chChart.Visible := False;
  chChartBarSeries1.Active := False;
  chChartBarSeries2.Active := False;
  chChartBarSeries3.Active := False;
  chChartBarSeries4.Active := False;
  chChartBarSeries5.Active := False;
  cbComboDiag1.Clear;
  cbComboDiag2.Clear;
  cbComboDiag3.Clear;
  cbComboDiag4.Clear;
  cbComboDiag5.Clear;
  clDiagBookmark.Clear;
  odOpenDialog.Filter := 'All formats|*.txt;*.odt;*.docx|' +
    'Text file|*.txt|Writer files|*.odt|' + 'Word files|*.docx|All files|*';
  odOpenDialog.DefaultExt := '.txt';
  odOpenDialog.FileName := '';
  if odOpenDialog.Execute then
    try
      if UTF8LowerCase(ExtractFileExt(odOpenDialog.FileName)) = '.txt' then
      begin
        meText.Lines.LoadFromFile(odOpenDialog.FileName);
        stFileName := odOpenDialog.FileName;
        blTextModSave := False;
        miFileSave.Enabled := blTextModSave;
        fmMain.Caption := 'WordStatix - ' + stFileName;
      end
      else if UTF8LowerCase(ExtractFileExt(odOpenDialog.FileName)) = '' then
      begin
        meText.Lines.LoadFromFile(odOpenDialog.FileName);
        stFileName := odOpenDialog.FileName;
        blTextModSave := False;
        miFileSave.Enabled := blTextModSave;
        fmMain.Caption := 'WordStatix - ' + stFileName;
      end
      else if UTF8LowerCase(ExtractFileExt(odOpenDialog.FileName)) = '.odt' then
      begin
        try
          Screen.Cursor := crHourGlass;
          Application.ProcessMessages;
          myZip := TUnZipper.Create;
          myList := TStringList.Create;
          stFileOrig := TStringList.Create;
          myList.Add('content.xml');
          myZip.OutputPath := GetTempDir;
          myZip.FileName := odOpenDialog.FileName;
          myZip.UnZipFiles(myList);
          stFileOrig.LoadFromFile(GetTempDir + DirectorySeparator + 'content.xml');
          stFileOrig.Text := StringReplace(stFileOrig.Text,
            '<text:note-citation>', ' [', [rfReplaceAll]);
          stFileOrig.Text := StringReplace(stFileOrig.Text,
            '</text:p></text:note-body></text:note>', ']', [rfReplaceAll]);
          stFileOrig.Text := StringReplace(stFileOrig.Text, '<text:note-body>',
            ': ', [rfReplaceAll]);
          stFileOrig.Text := StringReplace(stFileOrig.Text, '</text:h>',
            LineEnding + LineEnding, [rfReplaceAll]);
          stFileOrig.Text := StringReplace(stFileOrig.Text, '</text:p>',
            LineEnding, [rfReplaceAll]);
          stFileOrig.Text := StringReplace(stFileOrig.Text, '&apos;',
            #39, [rfReplaceAll]);
          meText.Text := CleanXML(stFileOrig.Text);
          blTextModSave := False;
          miFileSave.Enabled := blTextModSave;
          DeleteFileUTF8(GetTempDir + DirectorySeparator + 'content.xml');
        finally
          myZip.Free;
          myList.Free;
          stFileOrig.Free;
          Screen.Cursor := crDefault;
        end;
        stFileName := ExtractFileNameWithoutExt(odOpenDialog.FileName) + '.txt';
        fmMain.Caption := 'WordStatix - ' + stFileName;
      end
      else
      if UTF8LowerCase(ExtractFileExt(odOpenDialog.FileName)) = '.docx' then
      begin
        try
          Screen.Cursor := crHourGlass;
          Application.ProcessMessages;
          myZip := TUnZipper.Create;
          myList := TStringList.Create;
          stFileOrig := TStringList.Create;
          myZip.OutputPath := GetTempDir;
          myZip.FileName := odOpenDialog.FileName;
          myZip.UnZipAllFiles;
          if FileExistsUTF8(GetTempDir + 'word/document.xml') = True then
          begin
            stFileOrig.LoadFromFile(GetTempDir + 'word/document.xml');
            stFileOrig.Text :=
              StringReplace(stFileOrig.Text, '</w:p>', LineEnding, [rfReplaceAll]);
            i := 0;
            while Pos('<w:footnoteReference w:id="', stFileOrig.Text) > 0 do
            begin
              Inc(i);
              stFileOrig.Text :=
                StringReplace(stFileOrig.Text, '<w:footnoteReference w:id="',
                ' [' + IntToStr(i) + ']<', []);
            end;
            i := 0;
            while Pos('<w:endnoteReference w:id="', stFileOrig.Text) > 0 do
            begin
              Inc(i);
              stFileOrig.Text :=
                StringReplace(stFileOrig.Text, '<w:endnoteReference w:id="',
                ' [' + IntToStr(i) + ']<', []);
            end;
            meText.Text := meText.Text + CleanXML(stFileOrig.Text) + LineEnding;
          end;
          if FileExistsUTF8(GetTempDir + 'word/footnotes.xml') = True then
          begin
            stFileOrig.LoadFromFile(GetTempDir + 'word/footnotes.xml');
            stFileOrig.Text :=
              StringReplace(stFileOrig.Text, '</w:p>', LineEnding, [rfReplaceAll]);
            i := 0;
            while Pos('<w:footnoteRef/>', stFileOrig.Text) > 0 do
            begin
              Inc(i);
              stFileOrig.Text :=
                StringReplace(stFileOrig.Text, '<w:footnoteRef/>',
                '>[' + IntToStr(i) + '] ', []);
            end;
            meText.Text := meText.Text + CleanXML(stFileOrig.Text) + LineEnding;
          end;
          if FileExistsUTF8(GetTempDir + 'word/endnotes.xml') = True then
          begin
            stFileOrig.LoadFromFile(GetTempDir + 'word/endnotes.xml');
            stFileOrig.Text :=
              StringReplace(stFileOrig.Text, '</w:p>', LineEnding, [rfReplaceAll]);
            i := 0;
            while Pos('<w:endnoteRef/>', stFileOrig.Text) > 0 do
            begin
              Inc(i);
              stFileOrig.Text :=
                StringReplace(stFileOrig.Text, '<w:endnoteRef/>',
                '>[endnote: ' + IntToStr(i) + '] ', []);
            end;
            meText.Text := meText.Text + CleanXML(stFileOrig.Text) + LineEnding;
            blTextModSave := False;
            miFileSave.Enabled := blTextModSave;
          end;
          if DirectoryExistsUTF8(GetTempDir + 'word') = True then
          begin
            DeleteDirectory(GetTempDir + 'word', False);
          end;
        finally
          myZip.Free;
          myList.Free;
          stFileOrig.Free;
          Screen.Cursor := crDefault;
        end;
        stFileName := ExtractFileNameWithoutExt(odOpenDialog.FileName) + '.txt';
        fmMain.Caption := 'WordStatix - ' + stFileName;
      end;
      CreateBookmarks(True);
    except
      MessageDlg('It is not possible to open the selected file.',
        mtWarning, [mbOK], 0);
    end;
end;

procedure TfmMain.miFileSaveClick(Sender: TObject);
begin
  // Save file
  pcMain.ActivePage := tsFile;
  if stFileName <> '' then
    try
      meText.Lines.SaveToFile(stFileName);
      sbStatusBar.SimpleText := 'The text has been saved.';
      blTextModSave := False;
      miFileSave.Enabled := blTextModSave;
      fmMain.Caption := 'WordStatix - ' + stFileName;
    except
      MessageDlg('It is not possible to save the selected file.',
        mtWarning, [mbOK], 0);
    end
  else
  begin
    miFileSaveAsClick(nil);
  end;
end;

procedure TfmMain.miFileSaveAsClick(Sender: TObject);
begin
  // Salve as
  pcMain.ActivePage := tsFile;
  sdSaveDialog.Filter := 'Text file|*.txt|All files|*';
  sdSaveDialog.DefaultExt := '.txt';
  sdSaveDialog.FileName := '';
  if sdSaveDialog.Execute then
    try
      stFileName := sdSaveDialog.FileName;
      meText.Lines.SaveToFile(stFileName);
      sbStatusBar.SimpleText := 'The text has been saved.';
      blTextModSave := False;
      miFileSave.Enabled := blTextModSave;
      fmMain.Caption := 'WordStatix - ' + stFileName;
    except
      MessageDlg('It is not possible to save the selected file.',
        mtWarning, [mbOK], 0);
    end;
end;

procedure TfmMain.miFileOpenConcClick(Sender: TObject);
begin
  // Open concordance
  OpenConcordance;
end;

procedure TfmMain.miFileSaveConcClick(Sender: TObject);
begin
  // Save concordance
  SaveConcordance;
end;

procedure TfmMain.miFileSetBookmarkClick(Sender: TObject);
var
  stClip: string;
begin
  // Set bookmark
  // Use clipboard because othe rmethods
  // produce an unwanted vertical scrolling
  pcMain.ActivePage := tsFile;
  stClip := Clipboard.AsText;
  if meText.SelText = '' then
  begin
    if ((meText.SelStart = 0) or (UTF8Copy(meText.Text, meText.SelStart, 1) =
      LineEnding)) then
    begin
      Clipboard.AsText := '[[]] ';
    end
    else
    begin
      Clipboard.AsText := ' [[]] ';
    end;
    meText.PasteFromClipboard;
  end
  else
  begin
    Clipboard.AsText := '[[' + meText.SelText + ']]';
    Clipboard.AsText := StringReplace(Clipboard.AsText, ',', '', [rfReplaceAll]);
    Clipboard.AsText := StringReplace(Clipboard.AsText, ' ', '', [rfReplaceAll]);
    meText.PasteFromClipboard;
    CreateBookmarks(True);
  end;
  Clipboard.AsText := stClip;
end;

procedure TfmMain.miFileUpdBookmarkClick(Sender: TObject);
begin
  // Update bookmarks
  pcMain.ActivePage := tsFile;
  CreateBookmarks(True);
end;

procedure TfmMain.miFileExitClick(Sender: TObject);
begin
  // Quit
  pcMain.ActivePage := tsFile;
  Close;
end;

procedure TfmMain.miConcordanceCreateClick(Sender: TObject);
begin
  // Compile concordance
  pcMain.ActivePage := tsConcordance;
  if sgWordList.Visible = True then
    sgWordList.SetFocus;
  CreateConc(meText.Text);
  if sgWordList.Visible = True then
    sgWordList.SetFocus;
end;

procedure TfmMain.miConcodanceShowSelectedClick(Sender: TObject);
var
  i, iFirstVis: integer;
begin
  // Show only selected words
  pcMain.ActivePage := tsConcordance;
  if sgWordList.RowCount < 2 then
  begin
    Abort;
  end;
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  iFirstVis := 0;
  for i := 1 to sgWordList.RowCount - 1 do
  begin
    try
      if miConcodanceShowSelected.Checked = True then
      begin
        if sgWordList.Cells[2, i] = '0' then
        begin
          sgWordList.RowHeights[i] := 0;
        end
        else
        begin
          if iFirstVis = 0 then
          begin
            iFirstVis := i;
          end;
        end;
      end
      else
      begin
        sgWordList.RowHeights[i] := sgWordList.DefaultRowHeight;
        if iFirstVis = 0 then
        begin
          iFirstVis := i;
        end;
      end;
      sgWordList.SetFocus;
    finally
      Screen.Cursor := crDefault;
    end;
  end;
  if iFirstVis > 0 then
  begin
    sgWordList.Row := iFirstVis;
  end;
  if iFirstVis = 0 then
  begin
    lsContext.Clear;
  end
  else
  begin
    lsContext.Items.Text := CreateContext(sgWordList.Row, True);
    {$ifdef Win32}
    // Due to a bug in Windows...
    if lsContext.Items[0] = '' then
      lsContext.Items.Delete(0);
    {$endif}
    lsContext.ItemIndex := 0;
  end;
end;

procedure TfmMain.miConcordanceAddSkipClick(Sender: TObject);
begin
  // Add a word to skip
  pcMain.ActivePage := tsConcordance;
  AddSkipWord;
end;

procedure TfmMain.miConcordanceRemoveClick(Sender: TObject);
begin
  // Remove current word
  pcMain.ActivePage := tsConcordance;
  if sgWordList.RowCount > 1 then
  begin
    if sgWordList.RowHeights[sgWordList.Row] > 0 then
    begin
      if sgWordList.Cells[0, sgWordList.Row] <> '' then
      begin
        sgWordList.DeleteRow(sgWordList.Row);
        blGridConcMod := True;
      end;
    end;
    lsContext.Clear;
    if sgWordList.RowCount > 1 then
    begin
      if sgWordList.RowHeights[sgWordList.Row] > 0 then
      begin
        lsContext.Items.Text := CreateContext(sgWordList.Row, True);
        {$ifdef Win32}
        // Due to a bug in Windows...
        if lsContext.Items[0] = '' then
          lsContext.Items.Delete(0);
        {$endif}
        lsContext.ItemIndex := 0;
      end;
    end;
  end
  else
  begin
    MessageDlg('There is no word to remove.', mtWarning, [mbOK], 0);
  end;
end;

procedure TfmMain.miConcordanceJoinClick(Sender: TObject);
var
  i, iPos: integer;
  blSelection: boolean;
begin
  // Join selected words with current word
  if sgWordList.RowHeights[sgWordList.Row] = 0 then
  begin
    Abort;
  end;
  if sgWordList.RowCount < 2 then
  begin
    MessageDlg('No concordance has been created.', mtWarning, [mbOK], 0);
    Abort;
  end;
  blSelection := False;
  sgWordList.Cells[2, sgWordList.Row] := '0';
  for i := 1 to sgWordList.RowCount - 1 do
  begin
    if sgWordList.Cells[2, i] = '1' then
    begin
      blSelection := True;
      Break;
    end;
  end;
  if blSelection = False then
  begin
    MessageDlg('No words are selected.', mtWarning, [mbOK], 0);
    Abort;
  end;
  if MessageDlg('Associate all the recurrences of the selected words ' +
    'to those of the current word "' + sgWordList.Cells[0, sgWordList.Row] +
    '"?', mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
  begin
    Abort;
  end;
  for i := 1 to sgWordList.RowCount - 1 do
  begin
    if sgWordList.Cells[2, i] = '1' then
    begin
      sgWordList.Cells[1, sgWordList.Row] :=
        IntToStr(StrToInt(sgWordList.Cells[1, sgWordList.Row]) +
        StrToInt(sgWordList.Cells[1, i]));
      sgWordList.Cells[3, sgWordList.Row] :=
        sgWordList.Cells[3, sgWordList.Row] + sgWordList.Cells[3, i];
      sgWordList.Cells[4, sgWordList.Row] :=
        sgWordList.Cells[4, sgWordList.Row] + sgWordList.Cells[4, i];
    end;
  end;
  i := 1;
  iPos := sgWordList.Row;
  while i < sgWordList.RowCount do
  begin
    if sgWordList.Cells[2, i] = '1' then
    begin
      sgWordList.DeleteRow(i);
      if i < iPos then
      begin
        iPos := iPos - 1;
      end;
    end
    else
    begin
      Inc(i);
    end;
  end;
  blGridConcMod := True;
  sgWordList.Row := iPos;
  lsContext.Clear;
  lsContext.Items.Text := CreateContext(sgWordList.Row, True);
  {$ifdef Win32}
  // Due to a bug in Windows...
  if lsContext.Items[0] = '' then
    lsContext.Items.Delete(0);
  {$endif}
  lsContext.ItemIndex := 0;
end;

procedure TfmMain.miConcordanceRefreshGridClick(Sender: TObject);
begin
  // Refresh concordance grid
  pcMain.ActivePage := tsConcordance;
  if blGridConcMod = True then
  begin
    if MessageDlg('Some changes have been made to the concordance grid. ' +
      'Do you want to reject them and recreate the list of words?',
      mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
    begin
      Abort;
    end;
  end;
  CompileGrid;
end;

procedure TfmMain.miConcordanceOpenSkipClick(Sender: TObject);
begin
  // Open skip list
  pcMain.ActivePage := tsConcordance;
  odOpenDialog.Filter := 'Text file|*.txt|All files|*';
  odOpenDialog.DefaultExt := '.txt';
  odOpenDialog.FileName := '';
  if odOpenDialog.Execute then
    try
      meSkipList.Lines.LoadFromFile(odOpenDialog.FileName);
    except
      MessageDlg('It is not possible to open the skip list.',
        mtWarning, [mbOK], 0);
    end;
end;

procedure TfmMain.miConcordanceSaveSkipClick(Sender: TObject);
begin
  // Save skip list
  pcMain.ActivePage := tsConcordance;
  sdSaveDialog.Filter := 'Text file|*.txt|All files|*';
  sdSaveDialog.DefaultExt := '.txt';
  sdSaveDialog.FileName := '';
  if sdSaveDialog.Execute then
    try
      meSkipList.Lines.SaveToFile(sdSaveDialog.FileName);
    except
      MessageDlg('It is not possible to save the skip list.',
        mtWarning, [mbOK], 0);
    end;
end;

procedure TfmMain.miConcordanceDelContClick(Sender: TObject);
var
  slContList: TStringList;
  iIDList: integer;
begin
  // Selete selected recurrence
  if lsContext.Items.Count < 1 then
  begin
    MessageDlg('There is no recurrence to delete.', mtWarning, [mbOK], 0);
    Abort;
  end;
  if lsContext.ItemIndex < 0 then
  begin
    MessageDlg('There is no selected recurrence to delete.', mtWarning, [mbOK], 0);
    Abort;
  end;
  blGridConcMod := True;
  if lsContext.Items.Count = 1 then
  begin
    sgWordList.DeleteRow(sgWordList.Row);
    if sgWordList.RowCount > 1 then
    begin
      lsContext.Items.Text := CreateContext(sgWordList.Row, True);
      {$ifdef Win32}
      // Due to a bug in Windows...
      if lsContext.Items[0] = '' then
        lsContext.Items.Delete(0);
      {$endif}
      lsContext.ItemIndex := 0;
    end
    else
    begin
      lsContext.Clear;
    end;
  end
  else
    try
      slContList := TStringList.Create;
      slContList.CommaText := sgWordList.Cells[3, sgWordList.Row];
      slContList.Delete(lsContext.ItemIndex);
      sgWordList.Cells[3, sgWordList.Row] := slContList.CommaText;
      slContList.CommaText := sgWordList.Cells[4, sgWordList.Row];
      slContList.Delete(lsContext.ItemIndex);
      sgWordList.Cells[4, sgWordList.Row] := slContList.CommaText;
      sgWordList.Cells[1, sgWordList.Row] :=
        IntToStr(StrToInt(sgWordList.Cells[1, sgWordList.Row]) - 1);
      iIDList := lsContext.ItemIndex;
      // Sometimes it might gives error, so
      if lsContext.Items.Count > 0 then
      begin
        lsContext.Items.Delete(lsContext.ItemIndex);
      end;
      if iIDList < lsContext.Items.Count - 1 then
      begin
        lsContext.ItemIndex := iIDList;
      end
      else
      begin
        lsContext.ItemIndex := lsContext.Items.Count - 1;
      end;
    finally
      slContList.Free;
    end;
end;

procedure TfmMain.miConcordanceSaveRepClick(Sender: TObject);
begin
  // Export report to file
  pcMain.ActivePage := tsConcordance;
  if sgWordList.RowCount < 2 then
  begin
    MessageDlg('No concordance has been created.', mtWarning, [mbOK], 0);
    Abort;
  end;
  sdSaveDialog.Filter := 'Text file|*.txt|HTML file|*.html|All files|*';
  sdSaveDialog.DefaultExt := '.txt';
  sdSaveDialog.FileName := '';
  if sdSaveDialog.Execute then
    try
      SaveReport(sdSaveDialog.FileName);
      sbStatusBar.SimpleText := 'The report has been saved.';
    except
      MessageDlg('It is not possible to save the report.',
        mtWarning, [mbOK], 0);
    end;
end;

procedure TfmMain.miStatisticCreateClick(Sender: TObject);
begin
  // Create statistic
  pcMain.ActivePage := tsStatistic;
  CreateStatistic;
end;

procedure TfmMain.miStatisticSortNameClick(Sender: TObject);
var
  iPos1, iPos2, iCol: integer;
  slRow: TStringList;
begin
  // Sort by name the statistic
  if sgStatistic.RowCount = 0 then
  begin
    MessageDlg('No statistic available.', mtWarning, [mbOK], 0);
    Abort;
  end;
  try
    slRow := TStringList.Create;
    for iPos1 := 1 to sgStatistic.RowCount - 2 do
    begin
      for iPos2 := 1 to sgStatistic.RowCount - 3 do
      begin
        if cbCollatedSort.Checked = True then
        begin
          if UTF8CompareStrCollated(sgStatistic.Cells[0, iPos2],
            sgStatistic.Cells[0, iPos1]) > 0 then
          begin
            slRow.Clear;
            for iCol := 0 to sgStatistic.ColCount - 1 do
            begin
              slRow.Add(sgStatistic.Cells[iCol, iPos1]);
            end;
            sgStatistic.Rows[iPos1] := sgStatistic.Rows[iPos2];
            for iCol := 0 to slRow.Count - 1 do
            begin
              sgStatistic.Cells[iCol, iPos2] := slRow[iCol];
            end;
          end;
        end
        else
        begin
          if UTF8CompareStr(sgStatistic.Cells[0, iPos2],
            sgStatistic.Cells[0, iPos1]) > 0 then
          begin
            slRow.Clear;
            for iCol := 0 to sgStatistic.ColCount - 1 do
            begin
              slRow.Add(sgStatistic.Cells[iCol, iPos1]);
            end;
            sgStatistic.Rows[iPos1] := sgStatistic.Rows[iPos2];
            for iCol := 0 to slRow.Count - 1 do
            begin
              sgStatistic.Cells[iCol, iPos2] := slRow[iCol];
            end;
          end;
        end;
      end;
    end;
  finally
    slRow.Free;
  end;
end;

procedure TfmMain.miStatisticSortFreqClick(Sender: TObject);
var
  iPos1, iPos2, iCol: integer;
  slRow: TStringList;
begin
  // Sort by frequency the statistic
  if sgStatistic.RowCount = 0 then
  begin
    MessageDlg('No statistic available.', mtWarning, [mbOK], 0);
    Abort;
  end;
  try
    slRow := TStringList.Create;
    for iPos1 := 1 to sgStatistic.RowCount - 2 do
    begin
      for iPos2 := 1 to sgStatistic.RowCount - 3 do
      begin
        if StrToInt(sgStatistic.Cells[1, iPos2]) <
          StrToInt(sgStatistic.Cells[1, iPos1]) then
        begin
          slRow.Clear;
          for iCol := 0 to sgStatistic.ColCount - 1 do
          begin
            slRow.Add(sgStatistic.Cells[iCol, iPos1]);
          end;
          sgStatistic.Rows[iPos1] := sgStatistic.Rows[iPos2];
          for iCol := 0 to slRow.Count - 1 do
          begin
            sgStatistic.Cells[iCol, iPos2] := slRow[iCol];
          end;
        end
        else
        if StrToInt(sgStatistic.Cells[1, iPos2]) =
          StrToInt(sgStatistic.Cells[1, iPos1]) then
        begin
          if cbCollatedSort.Checked = True then
          begin
            if UTF8CompareStrCollated(sgStatistic.Cells[0, iPos2],
              sgStatistic.Cells[0, iPos1]) > 0 then
            begin
              slRow.Clear;
              for iCol := 0 to sgStatistic.ColCount - 1 do
              begin
                slRow.Add(sgStatistic.Cells[iCol, iPos1]);
              end;
              sgStatistic.Rows[iPos1] := sgStatistic.Rows[iPos2];
              for iCol := 0 to slRow.Count - 1 do
              begin
                sgStatistic.Cells[iCol, iPos2] := slRow[iCol];
              end;
            end;
          end
          else
          begin
            if UTF8CompareStr(sgStatistic.Cells[0, iPos2],
              sgStatistic.Cells[0, iPos1]) > 0 then
            begin
              slRow.Clear;
              for iCol := 0 to sgStatistic.ColCount - 1 do
              begin
                slRow.Add(sgStatistic.Cells[iCol, iPos1]);
              end;
              sgStatistic.Rows[iPos1] := sgStatistic.Rows[iPos2];
              for iCol := 0 to slRow.Count - 1 do
              begin
                sgStatistic.Cells[iCol, iPos2] := slRow[iCol];
              end;
            end;
          end;
        end;
      end;
    end;
  finally
    slRow.Free;
  end;
end;

procedure TfmMain.miStatisticSaveClick(Sender: TObject);
begin
  // Salve statistic
  pcMain.ActivePage := tsStatistic;
  if sgStatistic.RowCount = 0 then
  begin
    MessageDlg('There is no statistic to save.',
      mtWarning, [mbOK], 0);
    Abort;
  end;
  sdSaveDialog.Filter := 'CSV file|*.csv|All files|*';
  sdSaveDialog.DefaultExt := '.csv';
  sdSaveDialog.FileName := '';
  if sdSaveDialog.Execute then
    try
      sgStatistic.SaveToCSVFile(sdSaveDialog.FileName);
      sbStatusBar.SimpleText := 'The statistic has been saved.';
    except
      MessageDlg('It is not possible to save the selected file.',
        mtWarning, [mbOK], 0);
    end;
end;

procedure TfmMain.miDiagramAllWordNoBookClick(Sender: TObject);
begin
  // Create diagram of all words no bookmarks
  pcMain.ActivePage := tsDiagram;
  CreateDiagramAllWordsNoBook;
end;

procedure TfmMain.miDiagramAllWordsBookClick(Sender: TObject);
begin
  // Create diagram of all words in bookmarks
  pcMain.ActivePage := tsDiagram;
  CreateDiagramAllWordsBook;
end;

procedure TfmMain.miDiagramSingleWordsBookClick(Sender: TObject);
begin
  // Create diagram of single words in bookmarks
  pcMain.ActivePage := tsDiagram;
  CreateDiagramSingleWordsBook;
end;

procedure TfmMain.miDiagramSaveClick(Sender: TObject);
begin
  // Save diagram
  pcMain.ActivePage := tsDiagram;
  if chChart.Visible = False then
  begin
    MessageDlg('There is no diagram to save.', mtWarning, [mbOK], 0);
    Abort;
  end;
  sdSaveDialog.Filter := 'JPEG file|*.jpg|PNG file|*.png|All files|*';
  sdSaveDialog.DefaultExt := '.jpg';
  sdSaveDialog.FileName := '';
  if sdSaveDialog.Execute then
    try
      try
        chChart.Height := 1000;
        if UTF8LowerCase(ExtractFileExt(sdSaveDialog.FileName)) = '.png' then
        begin
          ChChart.SaveToFile(TPortableNetworkGraphic, sdSaveDialog.FileName);
        end
        else
        begin
          chChart.SaveToFile(TJpegImage, sdSaveDialog.FileName);
        end;
        sbStatusBar.SimpleText := 'The diagram has been saved.';
      except
        MessageDlg('It is not possible to save the diagram.',
          mtWarning, [mbOK], 0);
      end;
    finally
      chChart.Height := sbDiagram.Height - 10;
    end;
end;

procedure TfmMain.miDiagramShowValClick(Sender: TObject);
begin
  // Show values
  chChartBarSeries1.Marks.Visible := miDiagramShowVal.Checked;
  chChartBarSeries2.Marks.Visible := miDiagramShowVal.Checked;
  chChartBarSeries3.Marks.Visible := miDiagramShowVal.Checked;
  chChartBarSeries4.Marks.Visible := miDiagramShowVal.Checked;
  chChartBarSeries5.Marks.Visible := miDiagramShowVal.Checked;
  if ((miDiagramShowGrid.Checked = False) and (miDiagramShowVal.Checked = False)) then
  begin
    miDiagramShowGrid.Checked := True;
    miDiagramShowGridClick(nil);
  end;
end;

procedure TfmMain.miDiagramShowGridClick(Sender: TObject);
begin
  // Show axis
  chChart.AxisList[0].Visible := miDiagramShowGrid.Checked;
  chChart.AxisList[1].Grid.Visible := miDiagramShowGrid.Checked;
  if ((miDiagramShowGrid.Checked = False) and (miDiagramShowVal.Checked = False)) then
  begin
    miDiagramShowVal.Checked := True;
    miDiagramShowValClick(nil);
  end;
end;

procedure TfmMain.miDiaZoomInClick(Sender: TObject);
begin
  // Zoom in
  if chChart.Width < 9899 then
  begin
    chChart.Width := chChart.Width + 100;
  end;
end;

procedure TfmMain.miDiaZoomOutClick(Sender: TObject);
begin
  // Zoom out
  if chChart.Width > sbDiagram.Width + 97 then
  begin
    chChart.Width := chChart.Width - 100;
  end
  else
  begin
    chChart.Width := sbDiagram.Width - 3;
  end;
end;

procedure TfmMain.miDiaZoomNormalClick(Sender: TObject);
begin
  // Normal zoom
  chChart.Width := sbDiagram.Width - 3;
end;

procedure TfmMain.miManualClick(Sender: TObject);
begin
  // Show manual
  {$ifdef Linux}
  if OpenDocument(ExtractFileDir(Application.ExeName) + DirectorySeparator +
    'manual-wordstatix.pdf') = False then
  begin
    MessageDlg('No manual available in the folder of WordStatix. ' +
      'Download it from the website of the software:' + LineEnding +
      'https://sites.google.com/site/wordstatix.',
      mtInformation, [mbOK], 0);
  end;
  {$endif}
  {$ifdef Win32}
  if OpenDocument(ExtractFileDir(Application.ExeName) + DirectorySeparator +
    'manual-wordstatix.pdf') = False then
  begin
    MessageDlg('No manual available in the folder of WordStatix. ' +
      'Download it from the website of the software:' + LineEnding +
      'https://sites.google.com/site/wordstatix.',
      mtInformation, [mbOK], 0);
  end;
  {$endif}
  {$ifdef Darwin}
  if OpenDocument(StringReplace(ExtractFileDir(Application.ExeName),
    'wordstatix.app/Contents/MacOS', '', []) + DirectorySeparator +
    'manual-wordstatix.pdf') = False then
  begin
    MessageDlg('No manual available in the folder of WordStatix. ' +
      'Download it from the website of the software:' + LineEnding +
      'https://sites.google.com/site/wordstatix.',
      mtInformation, [mbOK], 0);
  end;
  {$endif}
end;

procedure TfmMain.miCopyrightFormClick(Sender: TObject);
begin
  // Copyright
  fmCopyright.ShowModal;
end;

// ********************************************************** //
// ******************* General procedures ******************* //
// ********************************************************** //

procedure TfmMain.CreateBookmarks(blSetCursor: boolean);
var
  iStart, iEnd, i, n: integer;
  stText, stOldBookmark: string;
  blCommasSpaces: boolean;
  slBookmark: TStringList;
begin
  // Set boomarks in the text
  stText := meText.Text;
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  try
    slBookmark := TStringList.Create;
    while UTF8Pos('[[', stText) > 0 do
    begin
      iStart := UTF8Pos('[[', stText) + 2;
      iEnd := UTF8Pos(']]', stText);
      if iEnd < iStart then
      begin
        lbBookmarks.Clear;
        if blSetCursor = True then
        begin
          Screen.Cursor := crDefault;
        end;
        MessageDlg('A bookmark is not open or closed properly ' +
          'with double square brackets.', mtWarning, [mbOK], 0);
        Exit;
      end
      else if iStart + 50 < iEnd then
      begin
        lbBookmarks.Clear;
        if blSetCursor = True then
        begin
          Screen.Cursor := crDefault;
        end;
        MessageDlg('A bookmark seems to be too long. ' +
          'Check the text to control that there are ' +
          'no unwanted double square brackets.', mtWarning, [mbOK], 0);
        Exit;
      end
      else
      if UTF8Pos(LineEnding, UTF8Copy(stText, iStart, iEnd - iStart)) > 0 then
      begin
        lbBookmarks.Clear;
        if blSetCursor = True then
        begin
          Screen.Cursor := crDefault;
        end;
        MessageDlg('A bookmark seems to be incorrect. ' +
          'Check the text to control that there are ' +
          'no unwanted double square brackets.', mtWarning, [mbOK], 0);
        Exit;
      end
      else if UTF8Copy(stText, iStart, iEnd - iStart) <> '' then
      begin
        slBookmark.Add(UTF8Copy(stText, iStart, iEnd - iStart));
      end;
      stText := UTF8Copy(stText, UTF8Pos(']]', stText) + 2, UTF8Length(stText));
    end;
    blCommasSpaces := False;
    for i := 0 to slBookmark.Count - 1 do
    begin
      if slBookmark[i] = '*' then
      begin
        slBookmark[i] := '.';
        meText.Text := StringReplace(meText.Text, '[[*]]', '[[.]]', []);
        blCommasSpaces := True;
      end;
      if ((UTF8Pos(',', slBookmark[i]) > 0) or
        (UTF8Pos(' ', slBookmark[i]) > 0)) then
      begin
        stOldBookmark := slBookmark[i];
        slBookmark[i] :=
          StringReplace(slBookmark[i], ',', '', [rfReplaceAll]);
        slBookmark[i] :=
          StringReplace(slBookmark[i], ' ', '_', [rfReplaceAll]);
        meText.Text := StringReplace(meText.Text, '[[' + stOldBookmark +
          ']]', '[[' + slBookmark[i] + ']]', [rfIgnoreCase]);
        blCommasSpaces := True;
      end;
    end;
    if slBookmark.Count > 1 then
    begin
      for i := 0 to slBookmark.Count - 2 do
      begin
        for n := i + 1 to slBookmark.Count - 1 do
        begin
          if UTF8LowerCase(slBookmark[i]) = UTF8LowerCase(slBookmark[n]) then
          begin
            lbBookmarks.Clear;
            if blSetCursor = True then
            begin
              Screen.Cursor := crDefault;
            end;
            MessageDlg('The bookmark "' + slBookmark[i] +
              '" is used more than once. ' +
              'Check the text and make all its recurrences unique.',
              mtWarning, [mbOK], 0);
            Exit;
          end;
        end;
      end;
    end;
    lbBookmarks.Clear;
    for i := 0 to slBookmark.Count - 1 do
    begin
      lbBookmarks.Items.Add(slBookmark[i]);
    end;
    if blCommasSpaces = True then
    begin
      if blSetCursor = True then
      begin
        Screen.Cursor := crDefault;
      end;
      MessageDlg('Bookmarks cannot contain commas, spaces or be an asterisk. ' +
        'The incorrect bookmarks have been changed accordingly.',
        mtWarning, [mbOK], 0);
    end;
  finally
    slBookmark.Free;
    if blSetCursor = True then
    begin
      Screen.Cursor := crDefault;
    end;
  end;
end;

function TfmMain.FindNextWord(blMessage: boolean): boolean;
begin
  // Find next word
  if edFind.Text <> '' then
  begin
    if UTF8Pos(UTF8LowerCase(edFind.Text), UTF8LowerCase(meText.Text),
      meText.SelStart + meText.SelLength + 1) > 0 then
    begin
      Result := True;
      meText.SelStart := UTF8Pos(UTF8LowerCase(edFind.Text),
        UTF8LowerCase(meText.Text), meText.SelStart + meText.SelLength + 1) - 1;
      meText.SelLength := UTF8Length(edFind.Text);
      meText.SetFocus;
    end
    else
    begin
      Result := False;
      if blMessage = True then
      begin
        MessageDlg('No more recurrences of the text to look for.',
          mtInformation, [mbOK], 0);
      end;
    end;
  end;
end;

procedure TfmMain.CreateConc(stStartText: string);
var
  slNoWord: TStringList;
  stWord, stBookmark: string;
  slText: TStringList;
  i, iTestNum: integer;
  flNewWord: boolean;
  flTmpFile: TextFile;

  // Create concordance
  procedure EvalStopProcess; inline;
  begin
    if blStopConcordance = True then
    begin
      sgWordList.RowCount := 1;
      lsContext.Clear;
      sgStatistic.RowCount := 0;
      sgStatistic.ColCount := 0;
      SetLength(arWordList, 0);
      SetLength(arWordTextual, 0);
      chChart.Visible := False;
      chChartBarSeries1.Active := False;
      chChartBarSeries2.Active := False;
      chChartBarSeries3.Active := False;
      chChartBarSeries4.Active := False;
      chChartBarSeries5.Active := False;
      cbComboDiag1.Clear;
      cbComboDiag2.Clear;
      cbComboDiag3.Clear;
      cbComboDiag4.Clear;
      cbComboDiag5.Clear;
      clDiagBookmark.Clear;
      sbStatusBar.SimpleText := 'Concordance interrupted.';
      EnableMenuItems;
      blConInProc := False;
      Screen.Cursor := crDefault;
      Abort;
    end;
  end;

begin
  if meText.Text = '' then
  begin
    pcMain.ActivePage := tsFile;
    MessageDlg('No text available.', mtWarning, [mbOK], 0);
    Abort;
  end
  else
  if sgWordList.RowCount > 1 then
  begin
    if MessageDlg('Replace the current concordance with a new one?',
      mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
      Abort;
  end;
  ttStart := Now;
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  sbStatusBar.SimpleText := 'Concordance in processing, please wait.';
  CreateBookmarks(False);
  blStopConcordance := False;
  blConInProc := True;
  DisableMenuItems;
  sgStatistic.RowCount := 0;
  sgStatistic.ColCount := 0;
  chChart.Visible := False;
  chChartBarSeries1.Active := False;
  chChartBarSeries2.Active := False;
  chChartBarSeries3.Active := False;
  chChartBarSeries4.Active := False;
  chChartBarSeries5.Active := False;
  cbComboDiag1.Clear;
  cbComboDiag2.Clear;
  cbComboDiag3.Clear;
  cbComboDiag4.Clear;
  cbComboDiag5.Clear;
  clDiagBookmark.Clear;
  sgWordList.Enabled := False;
  meSkipList.Text := CleanField(meSkipList.Text);
  edFltStart.Text := CleanField(edFltStart.Text);
  edFltEnd.Text := CleanField(edFltEnd.Text);
  SetLength(arWordList, 0);
  SetLength(arWordTextual, 0);
  iWordsUsed := 0;
  iWordsTotal := 0;
  stBookmark := '*';
  sgWordList.RowCount := 1;
  lsContext.Clear;
  stStartText := StringReplace(stStartText, #9, ' ', [rfReplaceAll]);
  stStartText := StringReplace(stStartText, #39, #39 + ' ', [rfReplaceAll]);
  stStartText := StringReplace(stStartText, #96, #39 + ' ', [rfReplaceAll]);
  stStartText := StringReplace(stStartText, '’', #39 + ' ', [rfReplaceAll]);
  while UTF8Pos('  ', stStartText) > 0 do
  begin
    stStartText := StringReplace(stStartText, '  ', ' ', [rfReplaceAll]);
  end;
  EvalStopProcess;
  try
    slText := TStringList.Create;
    slText.Delimiter := ' ';
    slText.QuoteChar := ' ';
    slText.DelimitedText := stStartText;
    sltext.SaveToFile(GetTempDir + DirectorySeparator + 'wrdstxtmp');
    iWordsStartTot := slText.Count;
  finally
    slText.Free;
  end;
  try
    slNoWord := TStringList.Create;
    slNoWord.CaseSensitive := False;
    slNoWord.CommaText := meSkipList.Text;
    sbStatusBar.SimpleText :=
      'Concordance in processing, press Ctrl + Shift + H to stop. ' +
      'Analyzed words: ' + FormatFloat('#,##0', iWordsTotal) +
      ' of ' + FormatFloat('#,##0', iWordsStartTot) + ' - Unique words found: ' +
      FormatFloat('#,##0', iWordsUsed) + '.';
    AssignFile(flTmpFile, GetTempDir + DirectorySeparator + 'wrdstxtmp');
    Reset(flTmpFile);
    while not EOF(flTmpFile) do
    begin
      ReadLn(flTmpFile, stWord);
      if UTF8Pos('[[', stWord) > 0 then
      begin
        stBookmark := UTF8Copy(stWord, UTF8Pos('[[', stWord) + 2,
          UTF8Pos(']]', stWord) - UTF8Pos('[[', stWord) - 2);
        if UTF8Length(stBookmark) > 20 then
          stBookmark := UTF8Copy(stBookmark, 1, 20) + '...';
        if stBookmark = '' then
        begin
          stBookmark := '*';
        end;
        Application.ProcessMessages;
        Continue;
      end;
      SetLength(arWordTextual, Length(arWordTextual) + 1);
      arWordTextual[Length(arWordTextual) - 1].stRecWordTextual := stWord;
      Inc(iWordsTotal);
      for i := 33 to 38 do
      begin
        stWord := StringReplace(stWord, Chr(i), '', [rfReplaceAll]);
      end;
      for i := 40 to 47 do
      begin
        stWord := StringReplace(stWord, Chr(i), '', [rfReplaceAll]);
      end;
      for i := 58 to 64 do
      begin
        stWord := StringReplace(stWord, Chr(i), '', [rfReplaceAll]);
      end;
      for i := 91 to 95 do
      begin
        stWord := StringReplace(stWord, Chr(i), '', [rfReplaceAll]);
      end;
      for i := 123 to 127 do
      begin
        stWord := StringReplace(stWord, Chr(i), '', [rfReplaceAll]);
      end;
      stWord := StringReplace(stWord, '«', '', [rfReplaceAll]);
      stWord := StringReplace(stWord, '»', '', [rfReplaceAll]);
      stWord := StringReplace(stWord, '“', '', [rfReplaceAll]);
      stWord := StringReplace(stWord, '”', '', [rfReplaceAll]);
      stWord := StringReplace(stWord, '–', '', [rfReplaceAll]);
      stWord := StringReplace(stWord, '  ', ' ', [rfReplaceAll]);
      if cbSkipNumbers.Checked = True then
      begin
        if TryStrToInt(stWord, iTestNum) = True then
        begin
          Application.ProcessMessages;
          SetLength(arWordTextual, Length(arWordTextual) - 1);
          Dec(iWordsTotal);
          Continue;
        end;
      end;
      if slNoWord.IndexOf(stWord) < 0 then
      begin
        if ((stWord <> ' ') and (stWord <> '') and (stWord <> #39) and
          (CheckFilter(stWord) = True)) then
        begin
          flNewWord := True;
          if Length(arWordList) > 0 then
          begin
            for i := 0 to Length(arWordList) - 1 do
            begin
              if UTF8LowerCase(stWord) = arWordList[i].stRecWord then
              begin
                flNewWord := False;
                Break;
              end;
            end;
          end;
          if flNewWord = True then
          begin
            SetLength(arWordList, Length(arWordList) + 1);
            arWordList[Length(arWordList) - 1].stRecWord := UTF8LowerCase(stWord);
            arWordList[Length(arWordList) - 1].iRecFreq := 1;
            arWordList[Length(arWordList) - 1].stRecWordPos :=
              arWordList[Length(arWordList) - 1].stRecWordPos +
              IntToStr(Length(arWordTextual) - 1) + ',';
            if stBookmark <> '' then
            begin
              arWordList[Length(arWordList) - 1].stRecBookPos :=
                arWordList[Length(arWordList) - 1].stRecBookPos +
                stBookmark + ',';
            end;
            Inc(iWordsUsed);
          end
          else
          begin
            arWordList[i].iRecFreq := arWordList[i].iRecFreq + 1;
            arWordList[i].stRecWordPos :=
              arWordList[i].stRecWordPos + IntToStr(Length(arWordTextual) - 1) + ',';
            if stBookmark <> '' then
            begin
              arWordList[i].stRecBookPos :=
                arWordList[i].stRecBookPos + stBookmark + ',';
            end;
          end;
        end;
      end;
      if iWordsStartTot > 10000 then
      begin
        if iWordsTotal mod 5000 = 0 then
        begin
          sbStatusBar.SimpleText :=
            'Concordance in processing, press Ctrl + Shift + H to stop. ' +
            'Analyzed words: ' + FormatFloat('#,##0', iWordsTotal) +
            ' of ' + FormatFloat('#,##0', iWordsStartTot) +
            ' - Unique words found: ' + FormatFloat('#,##0', iWordsUsed) + '.';
        end;
      end;
      Application.ProcessMessages;
      EvalStopProcess;
    end;
    CompileGrid;
  finally
    CloseFile(flTmpFile);
    slNoWord.Free;
    sgWordList.Enabled := True;
    EnableMenuItems;
    blConInProc := False;
    Screen.Cursor := crDefault;
  end;
  if FileExistsUTF8(GetTempDir + DirectorySeparator + 'wrdstxtmp') then
  begin
    DeleteFileUTF8(GetTempDir + DirectorySeparator + 'wrdstxtmp');
  end;
  blTextModConc := False;
  ttEnd := Now;
  sbStatusBar.SimpleText :=
    'Total words (without bookmarks): ' + FormatFloat('#,##0', iWordsTotal) +
    ' - Unique words: ' + FormatFloat('#,##0', iWordsUsed) +
    ' - Concordance done in ' + FormatDateTime('hh:nn:ss', ttEnd - ttStart) + '.';
  MessageDlg('The concordance has been created.' + LineEnding +
    LineEnding + 'Total words (without bookmarks): ' +
    FormatFloat('#,##0', iWordsTotal) + '.' + LineEnding + 'Unique words: ' +
    FormatFloat('#,##0', iWordsUsed) + '.' + LineEnding + 'Concordance done in ' +
    FormatDateTime('hh:nn:ss', ttEnd - ttStart) + '.',
    mtInformation, [mbOK], 0);
end;

function TfmMain.CreateContext(GridRow: integer; blSetCursor: boolean): string;
var
  stPos, stBookList, stBookmark, stContext, stDotsBf, stDotsAf: string;
  iPos, i: integer;
begin
  // Create context
  if blSetCursor = True then
  begin
    Screen.Cursor := crHourGlass;
  end;
  Application.ProcessMessages;
  Result := '';
  stPos := sgWordList.Cells[3, GridRow];
  stBookList := sgWordList.Cells[4, GridRow];
  stBookmark := '';
  try
    while Pos(',', stPos) > 0 do
    begin
      iPos := StrToInt(Copy(stPos, 1, Pos(',', stPos) - 1));
      stBookmark := Copy(stBookList, 1, Pos(',', stBookList) - 1);
      stContext := '';
      stDotsBf := '... ';
      stDotsAf := '...';
      for i := iPos - StrToInt(edWordsCont.Text)
        to iPos + StrToInt(edWordsCont.Text) do
      begin
        if i < 0 then
        begin
          stDotsBf := '';
          Continue;
        end;
        if i > Length(arWordTextual) - 1 then
        begin
          stDotsAf := '';
          Continue;
        end;
        if ((i = iPos - 1) or (i = iPos)) then
          stContext := stContext + arWordTextual[i].stRecWordTextual + ' • '
        else
          stContext := stContext + arWordTextual[i].stRecWordTextual + ' ';
      end;
      stContext := StringReplace(stContext, #39 + ' ', #39, [rfReplaceAll]);
      Result := Result + LineEnding + '[' + stBookmark + '] ' +
        stDotsBf + stContext + stDotsAf;
      stPos := Copy(stPos, Pos(',', stPos) + 1, Length(stPos));
      stBookList := Copy(stBookList, Pos(',', stBookList) + 1, Length(stBookList));
    end;
    Result := UTF8Copy(Result, 2, UTF8Length(Result));
  finally
    if blSetCursor = True then
    begin
      Screen.Cursor := crDefault;
    end;
  end;
end;

procedure TfmMain.AddSkipWord;
var
  WordsList: TStringList;
begin
  // Add selected word to skip list
  if sgWordList.RowCount > 1 then
  begin
    if sgWordList.RowHeights[sgWordList.Row] > 0 then
    begin
      if sgWordList.Cells[0, sgWordList.Row] <> '' then
      begin
        meSkipList.Text := CleanField(meSkipList.Text);
        try
          WordsList := TStringList.Create;
          WordsList.CommaText := meSkipList.Text;
          if WordsList.IndexOf(sgWordList.Cells[0, sgWordList.Row]) < 0 then
          begin
            meSkipList.Text :=
              meSkipList.Text + ', ' + sgWordList.Cells[0, sgWordList.Row];
            meSkipList.Text := CleanField(meSkipList.Text);
            sgWordList.DeleteRow(sgWordList.Row);
            blGridConcMod := True;
          end;
        finally
          WordsList.Free;
        end;
      end;
    end;
    lsContext.Clear;
    if sgWordList.RowCount > 1 then
    begin
      if sgWordList.RowHeights[sgWordList.Row] > 0 then
      begin
        lsContext.Items.Text := CreateContext(sgWordList.Row, True);
        {$ifdef Win32}
        // Due to a bug in Windows...
        if lsContext.Items[0] = '' then
          lsContext.Items.Delete(0);
        {$endif}
        lsContext.ItemIndex := 0;
      end;
    end;
  end
  else
  begin
    MessageDlg('There is no word to add to the skip list.', mtWarning, [mbOK], 0);
  end;
end;

procedure TfmMain.CompileGrid;
var
  i: integer;
begin
  // Compile grid of words
  if Length(arWordList) = 0 then
  begin
    sgWordList.RowCount := 1;
    lsContext.Clear;
    sbStatusBar.SimpleText :=
      'Total words (without bookmarks): ' + FormatFloat('#,##0', iWordsTotal) +
      ' - Unique words: ' + FormatFloat('#,##0', iWordsUsed) +
      ' - Concordance done in ' + FormatDateTime('hh:nn:ss',
      ttEnd - ttStart) + '.';
  end
  else
  begin
    try
      Screen.Cursor := crHourGlass;
      sgWordList.Enabled := False;
      sbStatusBar.SimpleText := 'List of words in sorting, please, wait.';
      Application.ProcessMessages;
      miConcodanceShowSelected.Checked := False;
      SortWordFreq(arWordList, rgSortBy.ItemIndex);
      sgWordList.RowCount := 1;
      sgWordList.RowCount := Length(arWordList) + 1;
      for i := 0 to Length(arWordList) - 1 do
      begin
        sgWordList.Cells[0, i + 1] := arWordList[i].stRecWord;
        sgWordList.Cells[1, i + 1] := IntToStr(arWordList[i].iRecFreq);
        sgWordList.Cells[2, i + 1] := '0';
        sgWordList.Cells[3, i + 1] := arWordList[i].stRecWordPos;
        sgWordList.Cells[4, i + 1] := arWordList[i].stRecBookPos;
        Application.ProcessMessages;
      end;
      lsContext.Clear;
      lsContext.Items.Text := CreateContext(sgWordList.Row, True);
    {$ifdef Win32}
      // Due to a bug in Windows...
      if lsContext.Items[0] = '' then
        lsContext.Items.Delete(0);
    {$endif}
      lsContext.ItemIndex := 0;
      blGridConcMod := False;
      sbStatusBar.SimpleText :=
        'Total words (without bookmarks): ' + FormatFloat('#,##0', iWordsTotal) +
        ' - Unique words: ' + FormatFloat('#,##0', iWordsUsed) +
        ' - Concordance done in ' + FormatDateTime('hh:nn:ss',
        ttEnd - ttStart) + '.';
    finally
      sgWordList.Enabled := True;
      Screen.Cursor := crDefault;
    end;
  end;
end;

function TfmMain.CheckFilter(stWord: string): boolean;
var
  slFltStart, slFltEnd: TStringList;
  i: integer;
  blStart, blEnd: boolean;
begin
  // Check the filtr on words
  Result := False;
  if ((edFltStart.Text = '') and (edFltEnd.Text = '')) then
  begin
    Result := True;
  end
  else
  begin
    blStart := False;
    blEnd := False;
    if edFltStart.Text <> '' then
    begin
      try
        slFltStart := TStringList.Create;
        slFltStart.CommaText := edFltStart.Text;
        for i := 0 to slFltStart.Count - 1 do
        begin
          if UTF8LowerCase(UTF8Copy(stWord, 1, UTF8Length(slFltStart[i]))) =
            UTF8LowerCase(slFltStart[i]) then
          begin
            blStart := True;
            Break;
          end;
        end;
      finally
        slFltStart.Free;
      end;
    end;
    if edFltEnd.Text <> '' then
    begin
      try
        slFltEnd := TStringList.Create;
        slFltEnd.CommaText := edFltEnd.Text;
        for i := 0 to slFltEnd.Count - 1 do
        begin
          if UTF8LowerCase(UTF8Copy(stWord, UTF8Length(stWord) -
            UTF8Length(slFltEnd[i]) + 1, UTF8Length(stWord))) =
            UTF8LowerCase(slFltEnd[i]) then
          begin
            blEnd := True;
            Break;
          end;
        end;
      finally
        slFltEnd.Free;
      end;
    end;
    if rgCondFlt.ItemIndex = 0 then
    begin
      if ((blStart = True) or (blEnd = True)) then
      begin
        Result := True;
      end
      else
      begin
        Result := False;
      end;
    end
    else
    begin
      if edFltStart.Text = '' then
      begin
        blStart := True;
      end;
      if edFltEnd.Text = '' then
      begin
        blEnd := True;
      end;
      if ((blStart = True) and (blEnd = True)) then
      begin
        Result := True;
      end
      else
      begin
        Result := False;
      end;
    end;
  end;
end;

procedure TfmMain.SortWordFreq(var arWordList: array of TRecWordList; flField: shortint);
var
  iPos1, iPos2: integer;
  rcTempRec: TRecWordList;
begin
  // Sort words
  for iPos1 := 0 to Length(arWordList) - 1 do
  begin
    for iPos2 := 0 to Length(arWordList) - 2 do
    begin
      if flField = 0 then
      begin
        if arWordList[iPos2].iRecFreq < arWordList[iPos2 + 1].iRecFreq then
        begin
          rcTempRec := arWordList[iPos2];
          arWordList[iPos2] := arWordList[iPos2 + 1];
          arWordList[iPos2 + 1] := rcTempRec;
        end
        else
        begin
          if arWordList[iPos2].iRecFreq = arWordList[iPos2 + 1].iRecFreq then
          begin
            if cbCollatedSort.Checked = True then
            begin
              if UTF8CompareStrCollated(arWordList[iPos2].stRecWord,
                arWordList[iPos2 + 1].stRecWord) > 0 then
              begin
                rcTempRec := arWordList[iPos2];
                arWordList[iPos2] := arWordList[iPos2 + 1];
                arWordList[iPos2 + 1] := rcTempRec;
              end;
            end
            else
            begin
              if UTF8CompareStr(arWordList[iPos2].stRecWord,
                arWordList[iPos2 + 1].stRecWord) > 0 then
              begin
                rcTempRec := arWordList[iPos2];
                arWordList[iPos2] := arWordList[iPos2 + 1];
                arWordList[iPos2 + 1] := rcTempRec;
              end;
            end;
          end;
        end;
      end
      else
      begin
        if cbCollatedSort.Checked = True then
        begin
          if UTF8CompareStrCollated(arWordList[iPos2].stRecWord,
            arWordList[iPos2 + 1].stRecWord) > 0 then
          begin
            rcTempRec := arWordList[iPos2];
            arWordList[iPos2] := arWordList[iPos2 + 1];
            arWordList[iPos2 + 1] := rcTempRec;
          end;
        end
        else
        begin
          if UTF8CompareStr(arWordList[iPos2].stRecWord,
            arWordList[iPos2 + 1].stRecWord) > 0 then
          begin
            rcTempRec := arWordList[iPos2];
            arWordList[iPos2] := arWordList[iPos2 + 1];
            arWordList[iPos2 + 1] := rcTempRec;
          end;
        end;
      end;
    end;
    Application.ProcessMessages;
  end;
end;

function TfmMain.CleanField(myField: string): string;
var
  myText: string;
  i: integer;
  WordsList, WordCurr: TStringList;
begin
  // Clean fields with commas by wrong inputs
  if myField <> '' then
  begin
    myText := myField;
    while ((UTF8Copy(myText, UTF8Length(myText), 1) = ',') or
        (UTF8Copy(myText, UTF8Length(myText), 1) = ' ')) do
    begin
      myText := UTF8Copy(myText, 1, UTF8Length(myText) - 1);
    end;
    myText := StringReplace(myText, ',', ', ', [rfReplaceAll]);
    myText := StringReplace(myText, ' ,', ',', [rfReplaceAll]);
    while Pos('  ', myText) > 0 do
    begin
      myText := StringReplace(myText, '  ', ' ', [rfReplaceAll]);
    end;
    while UTF8Copy(myText, 1, 1) = ' ' do
    begin
      myText := UTF8Copy(myText, 2, UTF8Length(myText));
    end;
    Result := myText;
    if myText <> '' then
    begin
      try
        WordsList := TStringList.Create;
        WordCurr := TStringList.Create;
        WordCurr.Text := StringReplace(myText, ', ', LineEnding, [rfReplaceAll]);
        WordCurr.Sort;
        for i := 0 to WordCurr.Count - 1 do
        begin
          if WordsList.IndexOf(WordCurr[i]) < 0 then
          begin
            if WordCurr[i] <> '' then
              WordsList.Add(WordCurr[i]);
          end;
        end;
        myText := WordsList.Text;
        Result := StringReplace(myText, LineEnding, ', ', [rfReplaceAll]);
        Result := UTF8Copy(Result, 1, UTF8Length(Result) - 2);
      finally
        WordsList.Free;
        WordCurr.Free;
      end;
    end;
  end
  else
  begin
    Result := '';
  end;
end;

procedure TfmMain.SaveConcordance;
var
  flConc: TextFile;
  i: integer;
  stConditions: string;
begin
  // Save concordance
  if meText.Text = '' then
  begin
    MessageDlg('No text available.', mtWarning, [mbOK], 0);
    Abort;
  end;
  if sgWordList.RowCount = 1 then
  begin
    MessageDlg('No concordance available.',
      mtWarning, [mbOK], 0);
    Abort;
  end;
  if blTextModConc = True then
  begin
    MessageDlg('The text or the settings of the software useful to create ' +
      'the concordance have been changed after the last concordance ' +
      'was processed. Recreate it to be able to save it along with its text.',
      mtWarning, [mbOK], 0);
    Abort;
  end;
  sdSaveDialog.Filter := 'WordStatix files|*.wsx|All files|*';
  sdSaveDialog.DefaultExt := '.wsx';
  sdSaveDialog.FileName := '';
  if sdSaveDialog.Execute then
  begin
    try
      try
        Screen.Cursor := crHourGlass;
        AssignFile(flConc, sdSaveDialog.FileName);
        Rewrite(flConc);
        WriteLn(flConc, '[WordStatixText]');
        for i := 0 to meText.Lines.Count - 1 do
        begin
          WriteLn(flConc, meText.Lines[i]);
        end;
        WriteLn(flConc, '[WordStatixList]');
        for i := Length(arWordList) - 1 downto 0 do
        begin
          WriteLn(flConc, arWordList[i].stRecWord + #9 +
            IntToStr(arWordList[i].iRecFreq) + #9 + arWordList[i].stRecWordPos +
            #9 + arWordList[i].stRecBookPos);
        end;
        WriteLn(flConc, '[WordStatixSeries]');
        for i := Length(arWordTextual) - 1 downto 0 do
        begin
          WriteLn(flConc, arWordTextual[i].stRecWordTextual);
        end;
        WriteLn(flConc, '[WordStatixConditions]');
        stConditions := '';
        stConditions := IntToStr(rgSortBy.ItemIndex) + #9;
        if cbCollatedSort.Checked = True then
        begin
          stConditions := stConditions + 'CST' + #9;
        end
        else
        begin
          stConditions := stConditions + 'CSF' + #9;
        end;
        if cbSkipNumbers.Checked = True then
        begin
          stConditions := stConditions + 'SNT' + #9;
        end
        else
        begin
          stConditions := stConditions + 'SNF' + #9;
        end;
        stConditions := stConditions + meSkipList.Lines[0] + #9;
        stConditions := stConditions + edFltStart.Text + #9;
        stConditions := stConditions + edFltEnd.Text + #9;
        stConditions := stConditions + IntToStr(rgCondFlt.ItemIndex) + #9;
        WriteLn(flConc, stCOnditions);
        WriteLn(flConc, '[WordStatixResults]');
        WriteLn(flConc, IntToStr(iWordsTotal) + #9 + IntToStr(iWordsUsed));
        sbStatusBar.SimpleText := 'The concordance has been saved.';
      finally
        CloseFile(flConc);
        Screen.Cursor := crDefault;
      end;
    except
      MessageDlg('It is not possibile to save the concordance.',
        mtWarning, [mbOK], 0);
    end;
  end;
end;

procedure TfmMain.OpenConcordance;
var
  flConc: TextFile;
  stLine: string;
  iSection: smallint;
begin
  // Open concordance
  pcMain.ActivePage := tsFile;
  if ((meText.Text <> '') and (blTextModSave = True)) then
  begin
    if MessageDlg('The current text has been changed but has not been saved. ' +
      'Reject the changes and open a new concordance with its text?',
      mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
    begin
      Abort;
    end;
  end;
  SetLength(arWordList, 0);
  SetLength(arWordTextual, 0);
  meText.Clear;
  lbBookmarks.Clear;
  edFind.Clear;
  edReplace.Clear;
  sgWordList.RowCount := 1;
  lsContext.Clear;
  edFltStart.Clear;
  edFltEnd.Clear;
  sgStatistic.RowCount := 0;
  sgStatistic.ColCount := 0;
  chChart.Visible := False;
  chChartBarSeries1.Active := False;
  chChartBarSeries2.Active := False;
  chChartBarSeries3.Active := False;
  chChartBarSeries4.Active := False;
  chChartBarSeries5.Active := False;
  cbComboDiag1.Clear;
  cbComboDiag2.Clear;
  cbComboDiag3.Clear;
  cbComboDiag4.Clear;
  cbComboDiag5.Clear;
  clDiagBookmark.Clear;
  odOpenDialog.Filter := 'WS concordances|*.wsx|All files|*';
  odOpenDialog.DefaultExt := '.wsx';
  odOpenDialog.FileName := '';
  if odOpenDialog.Execute then
  begin
    try
      Screen.Cursor := crHourGlass;
      ttStart := Now;
      iSection := -1;
      AssignFile(flConc, odOpenDialog.FileName);
      Reset(flConc);
      try
        while not EOF(flConc) do
        begin
          ReadLn(flConc, stLine);
          if stLine = '[WordStatixText]' then
          begin
            iSection := 0;
          end
          else if stLine = '[WordStatixList]' then
          begin
            iSection := 1;
          end
          else if stLine = '[WordStatixSeries]' then
          begin
            iSection := 2;
          end
          else if stLine = '[WordStatixConditions]' then
          begin
            iSection := 3;
          end
          else if stLine = '[WordStatixResults]' then
          begin
            iSection := 4;
          end
          else if iSection = 0 then
          begin
            meText.Lines.Add(stLine);
          end
          else if iSection = 1 then
          begin
            SetLength(arWordList, Length(arWordList) + 1);
            arWordList[Length(arWordList) - 1].stRecWord :=
              UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1);
            stLine := UTF8Copy(stLine, UTF8Pos(#9, stLine) + 1,
              UTF8Length(stLine));
            arWordList[Length(arWordList) - 1].iRecFreq :=
              StrToInt(UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1));
            stLine := UTF8Copy(stLine, UTF8Pos(#9, stLine) + 1,
              UTF8Length(stLine));
            arWordList[Length(arWordList) - 1].stRecWordPos :=
              UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1);
            stLine := UTF8Copy(stLine, UTF8Pos(#9, stLine) + 1,
              UTF8Length(stLine));
            arWordList[Length(arWordList) - 1].stRecBookPos :=
              UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1);
          end
          else if iSection = 2 then
          begin
            SetLength(arWordTextual, Length(arWordTextual) + 1);
            arWordTextual[Length(arWordTextual) - 1].stRecWordTextual := stLine;
          end
          else if iSection = 3 then
          begin
            rgSortBy.ItemIndex :=
              StrToInt(UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1));
            stLine := UTF8Copy(stLine, UTF8Pos(#9, stLine) + 1,
              UTF8Length(stLine));
            if UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1) = 'CST' then
            begin
              cbCollatedSort.Checked := True;
            end
            else if UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1) = 'CSF' then
            begin
              cbCollatedSort.Checked := False;
            end;
            stLine := UTF8Copy(stLine, UTF8Pos(#9, stLine) + 1,
              UTF8Length(stLine));
            if UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1) = 'SNT' then
            begin
              cbSkipNumbers.Checked := True;
            end
            else
            if UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1) = 'SNF' then
            begin
              cbSkipNumbers.Checked := False;
            end;
            stLine := UTF8Copy(stLine, UTF8Pos(#9, stLine) + 1,
              UTF8Length(stLine));
            meSkipList.Text := UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1);
            stLine := UTF8Copy(stLine, UTF8Pos(#9, stLine) + 1,
              UTF8Length(stLine));
            edFltStart.Text := UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1);
            stLine := UTF8Copy(stLine, UTF8Pos(#9, stLine) + 1,
              UTF8Length(stLine));
            edFltEnd.Text := UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1);
            stLine := UTF8Copy(stLine, UTF8Pos(#9, stLine) + 1,
              UTF8Length(stLine));
            rgCondFlt.ItemIndex :=
              StrToInt(UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1));
          end
          else if iSection = 4 then
          begin
            iWordsTotal := StrToInt(UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1));
            stLine := UTF8Copy(stLine, UTF8Pos(#9, stLine) + 1,
              UTF8Length(stLine));
            iWordsUsed := StrToInt(UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1));
          end
          else
          begin
            MessageDlg('The selected file was not created ' +
              'with WordStatix, and so cannot be opened.',
              mtWarning, [mbOK], 0);
            Exit;
          end;
        end;
      finally
        CloseFile(flConc);
        Screen.Cursor := crDefault;
      end;
      CompileGrid;
      blTextModSave := False;
      miFileSave.Enabled := blTextModSave;
      stFileName := ExtractFileNameWithoutExt(odOpenDialog.FileName) + '.txt';
      fmMain.Caption := 'WordStatix - ' + stFileName;
      pcMain.ActivePage := tsConcordance;
      ttEnd := Now;
      sbStatusBar.SimpleText :=
        'Total words (without bookmarks): ' + FormatFloat('#,##0', iWordsTotal) +
        ' - Unique words: ' + FormatFloat('#,##0', iWordsUsed) +
        ' - Concordance loaded in ' + FormatDateTime('hh:nn:ss',
        ttEnd - ttStart) + '.';
    except
      MessageDlg('It is not possibile to open the concordance.',
        mtWarning, [mbOK], 0);
    end;
  end;
end;

procedure TfmMain.SaveReport(stFileName: string);
var
  ConcFile: TextFile;
  i: integer;
  stTitle: string;
begin
  // Export concordance to file
  try
    Screen.Cursor := crHourGlass;
    Application.ProcessMessages;
    stTitle := InputBox('Title of concordance',
      'Insert the title of the concordance', '');
    if stTitle = '' then
      stTitle := 'Concordance made with WordStatix';
    AssignFile(ConcFile, stFileName);
    Rewrite(ConcFile);
    if UTF8LowerCase(ExtractFileExt(stFileName)) = '.html' then
    begin
      WriteLn(ConcFile, '<HTML>');
      WriteLn(ConcFile, '<HEAD>');
      WriteLn(ConcFile, '<meta http-equiv="Content-Type" ' +
        'content="text/html; charset=UTF-8">');
      WriteLn(ConcFile, '<TITLE>' + stTitle + '</TITLE>');
      WriteLn(ConcFile, ' </HEAD>');
      WriteLn(ConcFile, '<BODY>');
      WriteLn(ConcFile, '<H1>' + stTitle + '</H1>');
      for i := 1 to sgWordList.RowCount - 1 do
      begin
        if sgWordList.RowHeights[i] > 0 then
        begin
          WriteLn(ConcFile, '<H2>' + sgWordList.Cells[0, i] + '</H2>');
          WriteLn(ConcFile, '<b>' + sgWordList.Cells[1, i] + ' recurrences.</b>');
          WriteLn(ConcFile, '<BR>');
          WriteLn(ConcFile, StringReplace(CreateContext(i, False),
            LineEnding, '<BR>' + LineEnding, [rfReplaceAll]));
          WriteLn(ConcFile, '<BR>');
          WriteLn(ConcFile, '<BR>');
        end;
      end;
      WriteLn(ConcFile, '</BODY>');
      WriteLn(ConcFile, '</HTML>');
    end
    else
    begin
      WriteLn(ConcFile, stTitle);
      WriteLn(ConcFile, '');
      for i := 1 to sgWordList.RowCount - 1 do
      begin
        if sgWordList.RowHeights[i] > 0 then
        begin
          WriteLn(ConcFile, '*** ' + sgWordList.Cells[0, i] + ' ***');
          WriteLn(ConcFile, sgWordList.Cells[1, i] + ' recurrences.');
          WriteLn(ConcFile, '');
          WriteLn(ConcFile, CreateContext(i, False));
          WriteLn(ConcFile, '');
          WriteLn(ConcFile, '');
        end;
      end;
    end;
  finally
    CloseFile(ConcFile);
    Screen.Cursor := crDefault;
  end;
end;

function TfmMain.CleanXML(stXMLText: string): string;
var
  blTag: boolean;
  i: integer;
begin
  // Clean a text from XML/HTML tags
  // Do not use UTF8Lengt and UTF8Copy here
  blTag := False;
  Result := '';
  for i := 1 to Length(stXMLText) do
  begin
    if Copy(stXMLText, i, 1) = '<' then
    begin
      blTag := True;
    end
    else if Copy(stXMLText, i, 1) = '>' then
    begin
      blTag := False;
    end
    else if blTag = False then
    begin
      Result := Result + Copy(stXMLText, i, 1);
    end;
    Application.ProcessMessages;
  end;
  while Copy(Result, 1, 1) = LineEnding do
  begin
    Result := Copy(Result, 2, Length(Result));
  end;
end;

procedure TfmMain.CreateStatistic;
var
  i, n, iWord, iList, iTot: integer;
  slBookmark: TStringList;
begin
  // Create statistic
  if sgWordList.RowCount = 1 then
  begin
    pcMain.ActivePage := tsConcordance;
    MessageDlg('No concordance available.',
      mtWarning, [mbOK], 0);
    Abort;
  end;
  try
    Screen.Cursor := crHourGlass;
    Application.ProcessMessages;
    slBookmark := TStringList.Create;
    sgStatistic.RowCount := 1;
    sgStatistic.FixedRows := 1;
    sgStatistic.ColCount := lbBookmarks.Count + 3;
    sgStatistic.FixedCols := 1;
    cbComboDiag1.Clear;
    cbComboDiag2.Clear;
    cbComboDiag3.Clear;
    cbComboDiag4.Clear;
    cbComboDiag5.Clear;
    clDiagBookmark.Clear;
    for i := 1 to sgStatistic.ColCount - 1 do
    begin
      sgStatistic.ColWidths[i] := 70;
    end;
    sgStatistic.Cells[1, 0] := 'Total';
    sgStatistic.Cells[2, 0] := '*';
    clDiagBookmark.Items.Add('*');
    for i := 0 to lbBookmarks.Count - 1 do
    begin
      sgStatistic.Cells[i + 3, 0] := lbBookmarks.Items[i];
      clDiagBookmark.Items.Add(lbBookmarks.Items[i]);
    end;
    clDiagBookmark.CheckAll(cbChecked, False, False);
    for i := 1 to sgWordList.RowCount - 1 do
    begin
      if sgWordList.Cells[2, i] = '1' then
      begin
        sgStatistic.RowCount := sgStatistic.RowCount + 1;
        sgStatistic.Cells[0, sgStatistic.RowCount - 1] :=
          sgWordList.Cells[0, i];
        sgStatistic.Cells[1, sgStatistic.RowCount - 1] :=
          sgWordList.Cells[1, i];
        slBookmark.CommaText := sgWordList.Cells[4, i];
        cbComboDiag1.Items.Add(sgWordList.Cells[0, i]);
        cbComboDiag2.Items.Add(sgWordList.Cells[0, i]);
        cbComboDiag3.Items.Add(sgWordList.Cells[0, i]);
        cbComboDiag4.Items.Add(sgWordList.Cells[0, i]);
        cbComboDiag5.Items.Add(sgWordList.Cells[0, i]);
        for n := 2 to sgStatistic.ColCount - 1 do
        begin
          iWord := 0;
          for iList := 0 to slBookmark.Count - 1 do
          begin
            if UTF8LowerCase(slBookmark[iList]) =
              UTF8LowerCase(sgStatistic.Cells[n, 0]) then
              Inc(iWord);
          end;
          sgStatistic.Cells[n, sgStatistic.RowCount - 1] := IntToStr(iWord);
        end;
      end;
    end;
    if sgStatistic.RowCount = 1 then
    begin
      sgStatistic.RowCount := 0;
      sgStatistic.ColCount := 0;
      MessageDlg('No words are selected in the concordance gird.',
        mtWarning, [mbOK], 0);
      pcMain.ActivePage := tsConcordance;
      sgWordList.SetFocus;
    end
    else
    begin
      sgStatistic.RowCount := sgStatistic.RowCount + 1;
      sgStatistic.Cells[0, sgStatistic.RowCount - 1] := 'Total';
      for i := 1 to sgStatistic.ColCount - 1 do
      begin
        iTot := 0;
        for n := 1 to sgStatistic.RowCount - 2 do
        begin
          iTot := iTot + StrToInt(sgStatistic.Cells[i, n]);
        end;
        sgStatistic.Cells[i, sgStatistic.RowCount - 1] := IntToStr(iTot);
      end;
    end;
  finally
    Screen.Cursor := crDefault;
    slBookmark.Free;
  end;
end;

procedure TfmMain.CreateDiagramSingleWordsBook;
var
  iRowGridStat, iColGridStat, iNumDiag, iBottomAx, i: integer;
  clColor1, clColor2, clColor3, clColor4, clColor5: TColor;
  slCheckDouble: TStringList;
  blLastField, blSelBookmark: boolean;
begin
  // Create diagram of single words
  chChart.Visible := False;
  chChartBarSeries1.Active := False;
  chChartBarSeries2.Active := False;
  chChartBarSeries3.Active := False;
  chChartBarSeries4.Active := False;
  chChartBarSeries5.Active := False;
  if sgStatistic.RowCount = 0 then
  begin
    pcMain.ActivePage := tsStatistic;
    MessageDlg('No active statistic.',
      mtWarning, [mbOK], 0);
    Abort;
  end;
  if clDiagBookmark.Count = 0 then
  begin
    MessageDlg('No bookmark available.',
      mtWarning, [mbOK], 0);
    Abort;
  end;
  blSelBookmark := False;
  for i := 0 to clDiagBookmark.Count - 1 do
  begin
    if clDiagBookmark.Checked[i] = True then
    begin
      blSelBookmark := True;
    end;
  end;
  if blSelBookmark = False then
  begin
    MessageDlg('At least one bookmark must be selected.',
      mtWarning, [mbOK], 0);
    Abort;
  end;
  if cbComboDiag1.Text = '' then
  begin
    MessageDlg('No word selected in Word 1 field.',
      mtWarning, [mbOK], 0);
    Abort;
  end;
  try
    blLastField := False;
    slCheckDouble := TStringList.Create;
    slCheckDouble.Add(cbComboDiag1.Text);
    if cbComboDiag2.Text <> '' then
    begin
      slCheckDouble.Add(cbComboDiag2.Text);
    end
    else
    begin
      blLastField := True;
    end;
    if cbComboDiag3.Text <> '' then
    begin
      if blLastField = True then
      begin
        MessageDlg('The fields "Word" must be compiled contiguously.',
          mtWarning, [mbOK], 0);
        Abort;
      end
      else
      begin
        slCheckDouble.Add(cbComboDiag3.Text);
      end;
    end
    else
    begin
      blLastField := True;
    end;
    if cbComboDiag4.Text <> '' then
    begin
      if blLastField = True then
      begin
        MessageDlg('The fields "Word" must be compiled contiguously.',
          mtWarning, [mbOK], 0);
        Abort;
      end
      else
      begin
        slCheckDouble.Add(cbComboDiag4.Text);
      end;
    end
    else
    begin
      blLastField := True;
    end;
    if cbComboDiag5.Text <> '' then
    begin
      if blLastField = True then
      begin
        MessageDlg('The fields "Word" must be compiled contiguously.',
          mtWarning, [mbOK], 0);
        Abort;
      end
      else
      begin
        slCheckDouble.Add(cbComboDiag5.Text);
      end;
    end
    else
    begin
      blLastField := True;
    end;
    iNumDiag := slCheckDouble.Count;
    slCheckDouble.Sort;
    for i := 0 to slCheckDouble.Count - 2 do
    begin
      if slCheckDouble[i] = slCheckDouble[i + 1] then
      begin
        MessageDlg('The same word has been selected twice.',
          mtWarning, [mbOK], 0);
        Exit;
      end;
    end;
  finally
    slCheckDouble.Free;
  end;
  stDiagTitle := InputBox('Title of diagram', 'Insert the title of the diagram.',
    stDiagTitle);
  if stDiagTitle = '' then
  begin
    chChart.Title.Visible := False;
  end
  else
  begin
    chChart.Title.Text.Clear;
    chChart.Title.Text.Add(stDiagTitle);
    chChart.Title.Visible := True;
  end;
  clColor1 := $8080FF;
  clColor2 := $80FF80;
  clColor3 := $FFB980;
  clColor4 := $80FFFA;
  clColor5 := $E180FF;
  csChartSource1.DataPoints.Clear;
  csChartSource2.DataPoints.Clear;
  csChartSource3.DataPoints.Clear;
  csChartSource4.DataPoints.Clear;
  csChartSource5.DataPoints.Clear;
  chChartBarSeries1.Title := '';
  chChartBarSeries2.Title := '';
  chChartBarSeries3.Title := '';
  chChartBarSeries4.Title := '';
  chChartBarSeries5.Title := '';
  if cbComboDiag1.Text <> '' then
  begin
    for iRowGridStat := 1 to sgStatistic.RowCount - 1 do
    begin
      if sgStatistic.Cells[0, iRowGridStat] = cbComboDiag1.Text then
        Break;
    end;
    iBottomAx := 1;
    for iColGridStat := 2 to sgStatistic.ColCount - 1 do
    begin
      if clDiagBookmark.Checked[iColGridStat - 2] = True then
      begin
        csChartSource1.Add(
          iBottomAx,
          StrToInt(sgStatistic.Cells[iColGridStat, iRowGridStat]),
          sgStatistic.Cells[iColGridStat, 0],
          clColor1);
        Inc(iBottomAx);
      end;
    end;
    chChartBarSeries1.Title := cbComboDiag1.Text;
    chChartBarSeries1.SeriesColor := clColor1;
    chChartBarSeries1.Active := True;
  end;
  if cbComboDiag2.Text <> '' then
  begin
    for iRowGridStat := 1 to sgStatistic.RowCount - 1 do
    begin
      if sgStatistic.Cells[0, iRowGridStat] = cbComboDiag2.Text then
        Break;
    end;
    iBottomAx := 1;
    for iColGridStat := 2 to sgStatistic.ColCount - 1 do
    begin
      if clDiagBookmark.Checked[iColGridStat - 2] = True then
      begin
        csChartSource2.Add(
          iBottomAx,
          StrToInt(sgStatistic.Cells[iColGridStat, iRowGridStat]),
          sgStatistic.Cells[iColGridStat, 0],
          clColor2);
        Inc(iBottomAx);
      end;
    end;
    chChartBarSeries2.Title := cbComboDiag2.Text;
    chChartBarSeries2.SeriesColor := clColor2;
    chChartBarSeries2.Active := True;
  end;
  if cbComboDiag3.Text <> '' then
  begin
    for iRowGridStat := 1 to sgStatistic.RowCount - 1 do
    begin
      if sgStatistic.Cells[0, iRowGridStat] = cbComboDiag3.Text then
        Break;
    end;
    iBottomAx := 1;
    for iColGridStat := 2 to sgStatistic.ColCount - 1 do
    begin
      if clDiagBookmark.Checked[iColGridStat - 2] = True then
      begin
        csChartSource3.Add(
          iBottomAx,
          StrToInt(sgStatistic.Cells[iColGridStat, iRowGridStat]),
          sgStatistic.Cells[iColGridStat, 0],
          clColor3);
        Inc(iBottomAx);
      end;
    end;
    chChartBarSeries3.Title := cbComboDiag3.Text;
    chChartBarSeries3.SeriesColor := clColor3;
    chChartBarSeries3.Active := True;
  end;
  if cbComboDiag4.Text <> '' then
  begin
    for iRowGridStat := 1 to sgStatistic.RowCount - 1 do
    begin
      if sgStatistic.Cells[0, iRowGridStat] = cbComboDiag4.Text then
        Break;
    end;
    iBottomAx := 1;
    for iColGridStat := 2 to sgStatistic.ColCount - 1 do
    begin
      if clDiagBookmark.Checked[iColGridStat - 2] = True then
      begin
        csChartSource4.Add(
          iBottomAx,
          StrToInt(sgStatistic.Cells[iColGridStat, iRowGridStat]),
          sgStatistic.Cells[iColGridStat, 0],
          clColor4);
        Inc(iBottomAx);
      end;
    end;
    chChartBarSeries4.Title := cbComboDiag4.Text;
    chChartBarSeries4.SeriesColor := clColor4;
    chChartBarSeries4.Active := True;
  end;
  if cbComboDiag5.Text <> '' then
  begin
    for iRowGridStat := 1 to sgStatistic.RowCount - 1 do
    begin
      if sgStatistic.Cells[0, iRowGridStat] = cbComboDiag5.Text then
        Break;
    end;
    iBottomAx := 1;
    for iColGridStat := 2 to sgStatistic.ColCount - 1 do
    begin
      if clDiagBookmark.Checked[iColGridStat - 2] = True then
      begin
        csChartSource5.Add(
          iBottomAx,
          StrToInt(sgStatistic.Cells[iColGridStat, iRowGridStat]),
          sgStatistic.Cells[iColGridStat, 0],
          clColor5);
        Inc(iBottomAx);
      end;
    end;
    chChartBarSeries5.Title := cbComboDiag5.Text;
    chChartBarSeries5.SeriesColor := clColor5;
    chChartBarSeries5.Active := True;
  end;
  if iNumDiag = 1 then
  begin
    chChartBarSeries1.BarWidthPercent := 70;
    chChartBarSeries1.BarOffsetPercent := 0;
  end
  else if iNumDiag = 2 then
  begin
    chChartBarSeries1.BarWidthPercent := 35;
    chChartBarSeries1.BarOffsetPercent := -20;
    chChartBarSeries2.BarWidthPercent := 35;
    chChartBarSeries2.BarOffsetPercent := 20;
  end
  else if iNumDiag = 3 then
  begin
    chChartBarSeries1.BarWidthPercent := 25;
    chChartBarSeries1.BarOffsetPercent := -30;
    chChartBarSeries2.BarWidthPercent := 25;
    chChartBarSeries2.BarOffsetPercent := 0;
    chChartBarSeries3.BarWidthPercent := 25;
    chChartBarSeries3.BarOffsetPercent := 30;
  end
  else if iNumDiag = 4 then
  begin
    chChartBarSeries1.BarWidthPercent := 20;
    chChartBarSeries1.BarOffsetPercent := -34;
    chChartBarSeries2.BarWidthPercent := 20;
    chChartBarSeries2.BarOffsetPercent := -11;
    chChartBarSeries3.BarWidthPercent := 20;
    chChartBarSeries3.BarOffsetPercent := 11;
    chChartBarSeries4.BarWidthPercent := 20;
    chChartBarSeries4.BarOffsetPercent := 34;
  end
  else if iNumDiag = 5 then
  begin
    chChartBarSeries1.BarWidthPercent := 15;
    chChartBarSeries1.BarOffsetPercent := -34;
    chChartBarSeries2.BarWidthPercent := 15;
    chChartBarSeries2.BarOffsetPercent := -17;
    chChartBarSeries3.BarWidthPercent := 15;
    chChartBarSeries3.BarOffsetPercent := 0;
    chChartBarSeries4.BarWidthPercent := 15;
    chChartBarSeries4.BarOffsetPercent := 17;
    chChartBarSeries5.BarWidthPercent := 15;
    chChartBarSeries5.BarOffsetPercent := 34;
  end;
  chChart.Legend.Visible := True;
  chChart.Visible := True;
end;

procedure TfmMain.CreateDiagramAllWordsBook;
var
  iColGridStat, iBottomAx, i: integer;
  clColor1: TColor;
  blSelBookmark: boolean;
begin
  // Create diagram of all words
  chChart.Visible := False;
  chChartBarSeries1.Active := False;
  chChartBarSeries2.Active := False;
  chChartBarSeries3.Active := False;
  chChartBarSeries4.Active := False;
  chChartBarSeries5.Active := False;
  if sgStatistic.RowCount = 0 then
  begin
    pcMain.ActivePage := tsStatistic;
    MessageDlg('No active statistic.',
      mtWarning, [mbOK], 0);
    Abort;
  end;
  if clDiagBookmark.Count = 0 then
  begin
    MessageDlg('No bookmark available.',
      mtWarning, [mbOK], 0);
    Abort;
  end;
  blSelBookmark := False;
  for i := 0 to clDiagBookmark.Count - 1 do
  begin
    if clDiagBookmark.Checked[i] = True then
    begin
      blSelBookmark := True;
    end;
  end;
  if blSelBookmark = False then
  begin
    MessageDlg('At least one bookmark must be selected.',
      mtWarning, [mbOK], 0);
    Abort;
  end;
  stDiagTitle := InputBox('Title of diagram', 'Insert the title of the diagram.',
    stDiagTitle);
  if stDiagTitle = '' then
  begin
    chChart.Title.Visible := False;
  end
  else
  begin
    chChart.Title.Text.Clear;
    chChart.Title.Text.Add(stDiagTitle);
    chChart.Title.Visible := True;
  end;
  clColor1 := $8080FF;
  csChartSource1.DataPoints.Clear;
  chChartBarSeries1.Title := '';
  iBottomAx := 1;
  for iColGridStat := 2 to sgStatistic.ColCount - 1 do
  begin
    if clDiagBookmark.Checked[iColGridStat - 2] = True then
    begin
      csChartSource1.Add(
        iBottomAx,
        StrToInt(sgStatistic.Cells[iColGridStat, sgStatistic.RowCount - 1]),
        sgStatistic.Cells[iColGridStat, 0],
        clColor1);
      Inc(iBottomAx);
    end;
  end;
  chChartBarSeries1.SeriesColor := clColor1;
  chChartBarSeries1.BarWidthPercent := 70;
  chChartBarSeries1.BarOffsetPercent := 0;
  chChartBarSeries1.Active := True;
  chChart.Legend.Visible := False;
  chChart.Visible := True;
end;

procedure TfmMain.CreateDiagramAllWordsNoBook;
var
  iRowGridStat, iColGridStat, iBottomAx, iTot, i: integer;
  clColor1: TColor;
  blSelBookmark: boolean;
begin
  // Create diagram of all words with no bookmarks
  chChart.Visible := False;
  chChartBarSeries1.Active := False;
  chChartBarSeries2.Active := False;
  chChartBarSeries3.Active := False;
  chChartBarSeries4.Active := False;
  chChartBarSeries5.Active := False;
  if sgStatistic.RowCount = 0 then
  begin
    pcMain.ActivePage := tsStatistic;
    MessageDlg('No active statistic.',
      mtWarning, [mbOK], 0);
    Abort;
  end;
  if clDiagBookmark.Count = 0 then
  begin
    MessageDlg('No bookmark available.',
      mtWarning, [mbOK], 0);
    Abort;
  end;
  blSelBookmark := False;
  for i := 0 to clDiagBookmark.Count - 1 do
  begin
    if clDiagBookmark.Checked[i] = True then
    begin
      blSelBookmark := True;
    end;
  end;
  if blSelBookmark = False then
  begin
    MessageDlg('At least one bookmark must be selected.',
      mtWarning, [mbOK], 0);
    Abort;
  end;
  stDiagTitle := InputBox('Title of diagram', 'Insert the title of the diagram.',
    stDiagTitle);
  if stDiagTitle = '' then
  begin
    chChart.Title.Visible := False;
  end
  else
  begin
    chChart.Title.Text.Clear;
    chChart.Title.Text.Add(stDiagTitle);
    chChart.Title.Visible := True;
  end;
  clColor1 := $8080FF;
  csChartSource1.DataPoints.Clear;
  chChartBarSeries1.Title := '';
  iBottomAx := 1;
  for iRowGridStat := 1 to sgStatistic.RowCount - 2 do
  begin
    iTot := 0;
    for iColGridStat := 2 to sgStatistic.ColCount - 1 do
    begin
      if clDiagBookmark.Checked[iColGridStat - 2] = True then
      begin
        iTot := iTot + StrToInt(sgStatistic.Cells[iColGridStat, iRowGridStat]);
      end;
    end;
    csChartSource1.Add(
      iBottomAx,
      iTot,
      sgStatistic.Cells[0, iRowGridStat],
      clColor1);
    Inc(iBottomAx);
  end;
  chChartBarSeries1.SeriesColor := clColor1;
  chChartBarSeries1.BarWidthPercent := 70;
  chChartBarSeries1.BarOffsetPercent := 0;
  chChartBarSeries1.Active := True;
  chChart.Legend.Visible := False;
  chChart.Visible := True;
end;

procedure TfmMain.DisableMenuItems;
begin
  // Disable menu items
  miFileNew.Enabled := False;
  miFileOpen.Enabled := False;
  miFileSave.Enabled := False;
  miFileSaveAs.Enabled := False;
  miFileSetBookmark.Enabled := False;
  miFileUpdBookmark.Enabled := False;
  miConcordanceCreate.Enabled := False;
  miFileOpenConc.Enabled := False;
  miFileSaveConc.Enabled := False;
  miConcodanceShowSelected.Enabled := False;
  miConcordanceAddSkip.Enabled := False;
  miConcordanceDelCont.Enabled := False;
  miConcordanceRemove.Enabled := False;
  miConcordanceJoin.Enabled := False;
  miConcordanceRefreshGrid.Enabled := False;
  miConcordanceOpenSkip.Enabled := False;
  miConcordanceSaveSkip.Enabled := False;
  miConcordanceSaveRep.Enabled := False;
  miStatisticCreate.Enabled := False;
  miStatisticSortName.Enabled := False;
  miStatisticSortFreq.Enabled := False;
  miStatisticSave.Enabled := False;
  miDiagramSingleWordsBook.Enabled := False;
  miDiagramAllWordNoBook.Enabled := False;
  miDiagramAllWordsBook.Enabled := False;
  miDiagramShowVal.Enabled := False;
  miDiagramShowGrid.Enabled := False;
  miDiaZoomIn.Enabled := False;
  miDiaZoomOut.Enabled := False;
  miDiaZoomNormal.Enabled := False;
  miDiagramSave.Enabled := False;
  meSkipList.ReadOnly := True;
  edFltStart.ReadOnly := True;
  edFltEnd.ReadOnly := True;
  edLocate.ReadOnly := True;
  bnDeselConc.Enabled := False;
end;

procedure TfmMain.EnableMenuItems;
begin
  // Enable menu items
  miFileNew.Enabled := True;
  miFileOpen.Enabled := True;
  miFileSave.Enabled := blTextModSave;
  miFileSaveAs.Enabled := True;
  miFileSetBookmark.Enabled := True;
  miFileUpdBookmark.Enabled := True;
  miConcordanceCreate.Enabled := True;
  miFileOpenConc.Enabled := True;
  miFileSaveConc.Enabled := True;
  miConcodanceShowSelected.Enabled := True;
  miConcordanceAddSkip.Enabled := True;
  miConcordanceDelCont.Enabled := True;
  miConcordanceRemove.Enabled := True;
  miConcordanceJoin.Enabled := True;
  miConcordanceRefreshGrid.Enabled := True;
  miConcordanceOpenSkip.Enabled := True;
  miConcordanceSaveSkip.Enabled := True;
  miConcordanceSaveRep.Enabled := True;
  miStatisticCreate.Enabled := True;
  miStatisticSortName.Enabled := True;
  miStatisticSortFreq.Enabled := True;
  miStatisticSave.Enabled := True;
  miDiagramAllWordNoBook.Enabled := True;
  miDiagramAllWordsBook.Enabled := True;
  miDiagramSingleWordsBook.Enabled := True;
  miDiagramShowVal.Enabled := True;
  miDiagramShowGrid.Enabled := True;
  miDiaZoomIn.Enabled := True;
  miDiaZoomOut.Enabled := True;
  miDiaZoomNormal.Enabled := True;
  miDiagramSave.Enabled := True;
  meSkipList.ReadOnly := False;
  edFltStart.ReadOnly := False;
  edFltEnd.ReadOnly := False;
  edLocate.ReadOnly := False;
  bnDeselConc.Enabled := True;
end;

end.

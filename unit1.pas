// ***********************************************************************
// ***********************************************************************
// WordStatix 1.0
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
  StdCtrls, ExtCtrls, ComCtrls, Menus, LazUTF8, IniFiles, Clipbrd, zipper,
  LCLIntf;

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
    bnFindFirst: TButton;
    bnFindNext: TButton;
    bnReplace: TButton;
    bvWordsCont: TBevel;
    edFltStart: TEdit;
    edFltEnd: TEdit;
    edFind: TEdit;
    edReplace: TEdit;
    edWordsCont: TEdit;
    lbBookmarks: TListBox;
    lbReplace: TLabel;
    lbFind: TLabel;
    lbFltStart: TLabel;
    lbFltEnd: TLabel;
    lbWordsCont: TLabel;
    lbContext: TLabel;
    lbNoWord: TLabel;
    lsContext: TListBox;
    miManual: TMenuItem;
    miLine7: TMenuItem;
    miFileNew: TMenuItem;
    miStatisticSave: TMenuItem;
    miStatisticCreate: TMenuItem;
    miLine6: TMenuItem;
    miStatistic: TMenuItem;
    miFileSaveAs: TMenuItem;
    miLine5: TMenuItem;
    miConcordanceDelCont: TMenuItem;
    miLine1: TMenuItem;
    miFileUpdBookmark: TMenuItem;
    miFileSetBookmark: TMenuItem;
    miConcordanceOpenSkip: TMenuItem;
    miConcordanceSaveSkip: TMenuItem;
    miLine3: TMenuItem;
    miConcodanceShowSelected: TMenuItem;
    miFileSave: TMenuItem;
    miLine4: TMenuItem;
    miConcordanceSave: TMenuItem;
    miCopyrightForm: TMenuItem;
    miCopyright: TMenuItem;
    miLine2: TMenuItem;
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
    pnListBookmark: TPanel;
    pnTextBottom: TPanel;
    pcMain: TPageControl;
    rgSortBy: TRadioGroup;
    sdSaveDialog: TSaveDialog;
    sgWordList: TStringGrid;
    sbStatusBar: TStatusBar;
    sgStatistic: TStringGrid;
    spText: TSplitter;
    tsStatistic: TTabSheet;
    tsFile: TTabSheet;
    tsConcordance: TTabSheet;
    udWordsCont: TUpDown;
    procedure apAppPropException(Sender: TObject; E: Exception);
    procedure bnFindFirstClick(Sender: TObject);
    procedure bnFindNextClick(Sender: TObject);
    procedure bnReplaceClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure FormShow(Sender: TObject);
    procedure lbBookmarksMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: integer);
    procedure meSkipListExit(Sender: TObject);
    procedure meTextKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure meTextMouseWheel(Sender: TObject; Shift: TShiftState;
      WheelDelta: integer; MousePos: TPoint; var Handled: boolean);
    procedure miConcodanceShowSelectedClick(Sender: TObject);
    procedure miConcordanceAddSkipClick(Sender: TObject);
    procedure miConcordanceCreateClick(Sender: TObject);
    procedure miConcordanceOpenSkipClick(Sender: TObject);
    procedure miConcordanceSaveClick(Sender: TObject);
    procedure miConcordanceDelContClick(Sender: TObject);
    procedure miConcordanceSaveSkipClick(Sender: TObject);
    procedure miCopyrightFormClick(Sender: TObject);
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
    procedure pcMainChange(Sender: TObject);
    procedure rgSortByClick(Sender: TObject);
    procedure sgStatisticPrepareCanvas(Sender: TObject; aCol, aRow: integer;
      aState: TGridDrawState);
    procedure sgWordListKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure sgWordListSelection(Sender: TObject; aCol, aRow: integer);
    procedure udWordsContChanging(Sender: TObject; var AllowChange: boolean);
  private
    procedure CompileGrid;
    procedure CreateConc(stStartText: string);
    function CreateContext(GridRow: integer; blSetCursor: boolean): string;
    procedure AddSkipWord;
    function CheckFilter(stWord: string): boolean;
    function CleanField(myField: string): string;
    function CleanXML(stText: string): string;
    procedure CreateStatistic;
    procedure SaveConcordance(stFileName: string);
    procedure CreateBookmarks;
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
  ttStart, ttEnd: TTime;

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
  // Set color of the form and look
  fmMain.Color := fmMain.Color;
  if fmMain.Color <> clDefault then
  begin
    fmMain.sgWordList.FixedColor := fmMain.Color;
    lsContext.Color := fmMain.Color;
  end;
  sgWordList.FocusRectVisible := False;
  // Set home directory and data directories
  {$ifdef Linux}
  myHomeDir := GetEnvironmentVariable('HOME') + DirectorySeparator +
    '.config' + DirectorySeparator + 'wordstatix' + DirectorySeparator;
  myConfigFile := 'wordstatix';
  {$endif}
  {$ifdef Win32}
  myHomeDir := GetAppConfigDir(False);
  myConfigFile := 'wordstatix.ini';
  lbBookmarks.ScrollWidth := 0;
  lbBookmarks.Color := clWhite;
  lbWordsCont.Color := clWindow;
  {$endif}
  {$ifdef Darwin}
  myHomeDir := GetEnvironmentVariable('HOME') + DirectorySeparator +
    'Library' + DirectorySeparator + 'Preferences' + DirectorySeparator;
  myConfigFile := 'wordstatix.plist';
  {$endif}
  if DirectoryExists(myHomeDir) = False then
  begin
    CreateDirUTF8(myHomeDir);
  end;
  if FileExists(myHomeDir + 'skipwords') then
  begin
    meSkipList.Lines.LoadFromFile(myHomeDir + 'skipwords');
  end;
  if FileExists(myHomeDir + myConfigFile) then
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
      if MyIni.ReadString('wordstatix', 'wordscont', '') <> '' then
      begin
        edWordsCont.Text := MyIni.ReadString('wordstatix', 'wordscont', '6');
      end;
    finally
      MyIni.Free;
    end;
end;

procedure TfmMain.FormClose(Sender: TObject; var CloseAction: TCloseAction);
var
  MyIni: TIniFile;
begin
  // Save skip words
  meSkipList.Lines.SaveToFile(myHomeDir + 'skipwords');
  // Save main form dimensions and other elements to ini file
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

procedure TfmMain.rgSortByClick(Sender: TObject);
begin
  // Compile the grid with the words
  CompileGrid;
end;

procedure TfmMain.pcMainChange(Sender: TObject);
begin
  // Selecting the tsConcordance the sdGrid do not accept focus, so...
  if pcMain.ActivePage = tsConcordance then
    tsConcordance.SetFocus;
end;

procedure TfmMain.sgStatisticPrepareCanvas(Sender: TObject;
  aCol, aRow: integer; aState: TGridDrawState);
begin
  // Color the last row in Statistic
  if aRow = sgStatistic.RowCount - 1 then
  begin
    sgStatistic.Canvas.Font.Color := clWhite;
    sgStatistic.Canvas.Brush.Color := clDkGray; //sgStatistic.FixedColor;
  end;
end;

procedure TfmMain.sgWordListKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  // Select and deselect the word
  if key = 32 then
  begin
    if sgWordList.Cells[2, sgWordList.Row] = '0' then
    begin
      sgWordList.Cells[2, sgWordList.Row] := '1';
    end
    else
    begin
      sgWordList.Cells[2, sgWordList.Row] := '0';
    end;
  end;
end;

procedure TfmMain.sgWordListSelection(Sender: TObject; aCol, aRow: integer);
begin
  // Create the list of context
  if sgWordList.RowCount > 1 then
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

procedure TfmMain.udWordsContChanging(Sender: TObject; var AllowChange: boolean);
begin
  // Create the list of context
  if sgWordList.RowCount > 1 then
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

procedure TfmMain.meSkipListExit(Sender: TObject);
begin
  //Clean skip words on exit
  meSkipList.Text := CleanField(meSkipList.Text);
end;

procedure TfmMain.meTextKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  // Change Zoom with keys
  if ((Key = 187) and (Shift = [ssCtrl])) then
  begin
    if meText.Font.Size < 48 then
      meText.Font.Size := meText.Font.Size + 1;
  end
  else if ((Key = 189) and (Shift = [ssCtrl])) then
  begin
    if meText.Font.Size > 6 then
      meText.Font.Size := meText.Font.Size - 1;
  end
  else if ((Key = Ord('0')) and (Shift = [ssCtrl])) then
  begin
    meText.Font.Size := 14;
  end
  else if ((Key = Ord('K')) and (Shift = [ssCtrl, ssShift])) then
  begin
    if MessageDlg('Compact alla paragraphs not separated by an empty line' +
      LineEnding + 'and remove double spaces in the text?', mtConfirmation,
      [mbOK, mbCancel], 0) = mrCancel then
      Abort;
    meText.Text := StringReplace(meText.Text, LineEnding + LineEnding,
      #2, [rfReplaceAll]);
    meText.Text := StringReplace(meText.Text, LineEnding, ' ', [rfReplaceAll]);
    meText.Text := StringReplace(meText.Text, #2, LineEnding +
      LineEnding, [rfReplaceAll]);
    meText.Text := StringReplace(meText.Text, '  ', ' ', [rfReplaceAll]);
  end;
end;

procedure TfmMain.bnFindFirstClick(Sender: TObject);
begin
  // Find first
  if edFind.Text <> '' then
  begin
    if UTF8Pos(edFind.Text, meText.Text, 1) > 0 then
    begin
      meText.SelStart := UTF8Pos(edFind.Text, meText.Text, 1) - 1;
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
  if edFind.Text <> '' then
  begin
    if UTF8Pos(edFind.Text, meText.Text, meText.SelStart +
      meText.SelLength + 1) > 0 then
    begin
      meText.SelStart := UTF8Pos(edFind.Text, meText.Text, meText.SelStart +
        meText.SelLength + 1) - 1;
      meText.SelLength := UTF8Length(edFind.Text);
      meText.SetFocus;
    end
    else
    begin
      MessageDlg('No more recurrences of the text to look for.',
        mtInformation, [mbOK], 0);
    end;
  end;
end;

procedure TfmMain.bnReplaceClick(Sender: TObject);
var
  stClip: string;
begin
  // Replace selection
  // Use clipboard because other methods
  // produce an unwanted vertical scrolling
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

// ********************************************************** //
// ****************** Events of menu items ****************** //
// ********************************************************** //

procedure TfmMain.miFileNewClick(Sender: TObject);
begin
  // New
  if meText.Text <> '' then
  begin
    if MessageDlg('Create a new text and a new concordance?',
      mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
      Abort;
  end;
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
end;

procedure TfmMain.miFileOpenClick(Sender: TObject);
var
  myZip: TUnZipper;
  myList, stFileOrig: TStringList;
begin
  // Open file
  pcMain.ActivePage := tsFile;
  if meText.Text <> '' then
  begin
    if MessageDlg('Open a new text replacing the existing one?',
      mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
      Abort;
  end;
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
  odOpenDialog.Filter := 'Text file|*.txt|Writer file|*.odt|All files|*';
  odOpenDialog.DefaultExt := '.txt';
  odOpenDialog.FileName := '';
  if odOpenDialog.Execute then
    try
      if UTF8LowerCase(ExtractFileExt(odOpenDialog.FileName)) = '.txt' then
      begin
        meText.Lines.LoadFromFile(odOpenDialog.FileName);
        stFileName := odOpenDialog.FileName;
        fmMain.Caption := 'WordStatix - ' + stFileName;
      end
      else if UTF8LowerCase(ExtractFileExt(odOpenDialog.FileName)) = '' then
      begin
        meText.Lines.LoadFromFile(odOpenDialog.FileName);
        stFileName := odOpenDialog.FileName;
        fmMain.Caption := 'WordStatix - ' + stFileName;
      end
      else if UTF8LowerCase(ExtractFileExt(odOpenDialog.FileName)) = '.odt' then
        try
          myZip := TUnZipper.Create;
          myList := TStringList.Create;
          stFileOrig := TStringList.Create;
          myList.Add('content.xml');
          myZip.OutputPath := GetTempDir;
          myZip.FileName := odOpenDialog.FileName;
          myZip.UnZipFiles(myList);
          stFileOrig.LoadFromFile(GetTempDir + DirectorySeparator + 'content.xml');
          DeleteFile(GetTempDir + DirectorySeparator + 'content.xml');
          stFileOrig.Text := StringReplace(stFileOrig.Text,
            '<text:note-citation>', '[', [rfReplaceAll]);
          stFileOrig.Text := StringReplace(stFileOrig.Text,
            '</text:p></text:note-body></text:note>', ']', [rfReplaceAll]);
          stFileOrig.Text := StringReplace(stFileOrig.Text, '<text:note-body>',
            ': ', [rfReplaceAll]);
          stFileOrig.Text := StringReplace(stFileOrig.Text, '</text:h>',
            LineEnding + LineEnding, [rfReplaceAll]);
          stFileOrig.Text := StringReplace(stFileOrig.Text, '</text:p>',
            LineEnding, [rfReplaceAll]);
          meText.Text := CleanXML(stFileOrig.Text);
        finally
          myZip.Free;
          myList.Free;
          stFileOrig.Free;
        end;
      CreateBookmarks;
    except
      MessageDlg('It'' not possible to open the selected file.',
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
      fmMain.Caption := 'WordStatix - ' + stFileName;
    except
      MessageDlg('It'' not possible to save the selected file.',
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
      fmMain.Caption := 'WordStatix - ' + stFileName;
    except
      MessageDlg('It'' not possible to save the selected file.',
        mtWarning, [mbOK], 0);
    end;
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
    Clipboard.AsText := '[[]]';
    meText.PasteFromClipboard;
  end
  else
  begin
    Clipboard.AsText := '[[' + meText.SelText + ']]';
    Clipboard.AsText := StringReplace(Clipboard.AsText, ',', '', [rfReplaceAll]);
    Clipboard.AsText := StringReplace(Clipboard.AsText, ' ', '', [rfReplaceAll]);
    meText.PasteFromClipboard;
  end;
  Clipboard.AsText := stClip;
  CreateBookmarks;
end;

procedure TfmMain.miFileUpdBookmarkClick(Sender: TObject);
begin
  // Update bookmarks
  pcMain.ActivePage := tsFile;
  CreateBookmarks;
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
  CreateConc(meText.Text);
  if sgWordList.Visible = True then
    sgWordList.SetFocus;
end;

procedure TfmMain.miConcodanceShowSelectedClick(Sender: TObject);
var
  i: integer;
begin
  // Show only selected words
  pcMain.ActivePage := tsConcordance;
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  for i := 1 to sgWordList.RowCount - 1 do
    try
      if miConcodanceShowSelected.Checked = True then
      begin
        if sgWordList.Cells[2, i] = '0' then
        begin
          sgWordList.RowHeights[i] := 0;
        end;
      end
      else
      begin
        sgWordList.RowHeights[i] := sgWordList.DefaultRowHeight;
      end;
    finally
      Screen.Cursor := crDefault;
    end;
end;

procedure TfmMain.miConcordanceAddSkipClick(Sender: TObject);
begin
  // Add a word to skip
  pcMain.ActivePage := tsConcordance;
  AddSkipWord;
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
      MessageDlg('It'' not possible to open the skip list.',
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
      MessageDlg('It'' not possible to save the skip list.',
        mtWarning, [mbOK], 0);
    end;
end;

procedure TfmMain.miConcordanceDelContClick(Sender: TObject);
var
  slContList: TStringList;
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
  if MessageDlg('Delete selected recurrence in context list?',
    mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
  begin
    Abort;
  end;
  if lsContext.Items.Count = 1 then
  begin
    sgWordList.DeleteRow(sgWordList.Row);
    lsContext.Clear;
    if sgWordList.RowCount > 1 then
    begin
      lsContext.Items.Text := CreateContext(sgWordList.Row, True);
      {$ifdef Win32}
      // Due to a bug in Windows...
      if lsContext.Items[0] = '' then
        lsContext.Items.Delete(0);
      {$endif}
      lsContext.ItemIndex := 0;
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
    finally
      slContList.Free;
    end;
  lsContext.Items.Delete(lsContext.ItemIndex);
end;

procedure TfmMain.miConcordanceSaveClick(Sender: TObject);
begin
  // Export concordance to file
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
      SaveConcordance(sdSaveDialog.FileName);
    except
      MessageDlg('It'' not possible to save the concordance.',
        mtWarning, [mbOK], 0);
    end;
end;

procedure TfmMain.miStatisticCreateClick(Sender: TObject);
begin
  // Create statistic
  pcMain.ActivePage := tsStatistic;
  CreateStatistic;
end;

procedure TfmMain.miStatisticSaveClick(Sender: TObject);
begin
  // Salve statistic
  pcMain.ActivePage := tsStatistic;
  if sgStatistic.RowCount < 2 then
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
    except
      MessageDlg('It'' not possible to save the selected file.',
        mtWarning, [mbOK], 0);
    end;
end;

procedure TfmMain.miManualClick(Sender: TObject);
begin
  // Show manual
  {$ifdef Linux}
  if OpenDocument('/opt/wordstatix/manual-wordstatix.pdf') = False then
  begin
    MessageDlg('No manual available. Download it from' + LineEnding +
      'the website of the software:' + LineEnding +
      'https://sites.google.com/site/wordstatix',
      mtInformation, [mbOK], 0);
  end;
  {$endif}
  {$ifdef Win32}
  if OpenDocument('manual-wordstatix.pdf') = False then
  begin
    MessageDlg('No manual available. Download it from' + LineEnding +
      'the website of the software:' + LineEnding +
      'https://sites.google.com/site/wordstatix',
      mtInformation, [mbOK], 0);
  end;
  {$endif}
  {$ifdef Darwin}
  if OpenDocument('manual-wordstatix.pdf') = False then
  begin
    MessageDlg('No manual available. Download it from' + LineEnding +
      'the website of the software:' + LineEnding +
      'https://sites.google.com/site/wordstatix',
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

procedure TfmMain.CreateBookmarks;
var
  iStart, iEnd, i: integer;
  stText, oldBookmark: string;
  blCommasSpaces: boolean;
begin
  // Set boomarks in the text
  lbBookmarks.Clear;
  stText := meText.Text;
  while UTF8Pos('[[', stText) > 0 do
  begin
    iStart := UTF8Pos('[[', stText) + 2;
    iEnd := UTF8Pos(']]', stText);
    if iStart > iEnd + 50 then
    begin
      MessageDlg('A bookmark seems to be too long;' + LineEnding +
        'check the text to control that there are' + LineEnding +
        'no unwanted double square brakets.', mtWarning, [mbOK], 0);
      Abort;
    end
    else if ((UTF8Copy(stText, iStart, iEnd - iStart) <> '') and
      (UTF8Copy(stText, iStart, iEnd - iStart) <> ' ')) then
    begin
      lbBookmarks.Items.Add(UTF8Copy(stText, iStart, iEnd - iStart));
    end;
    stText := UTF8Copy(stText, UTF8Pos(']]', stText) + 2, UTF8Length(stText));
  end;
  blCommasSpaces := False;
  for i := 0 to lbBookmarks.Items.Count - 1 do
  begin
    if ((UTF8Pos(',', lbBookmarks.Items[i]) > 0) or
      (UTF8Pos(' ', lbBookmarks.Items[i]) > 0)) then
    begin
      oldBookmark := lbBookmarks.Items[i];
      lbBookmarks.Items[i] := StringReplace(lbBookmarks.Items[i],
        ',', '', [rfReplaceAll]);
      lbBookmarks.Items[i] := StringReplace(lbBookmarks.Items[i],
        ' ', '_', [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, '[[' + oldBookmark +
        ']]', '[[' + lbBookmarks.Items[i] + ']]', [rfIgnoreCase]);
      blCommasSpaces := True;
    end;
  end;
  if blCommasSpaces = True then
  begin
    MessageDlg('Bookmarks cannot contain commas or spaces,' +
      LineEnding + 'so they have been removed.', mtWarning, [mbOK], 0);
  end;
end;

procedure TfmMain.CreateConc(stStartText: string);
var
  slNoWord: TStringList;
  stText, stWord, stBookmark: string;
  i: integer;
  flNewWord: boolean;
begin
  // Create concordance
  if meText.Text = '' then
  begin
    MessageDlg('No text available.', mtWarning, [mbOK], 0);
    Abort;
  end;
  sgWordList.Enabled := False;
  ttStart := Now;
  CreateBookmarks;
  SetLength(arWordList, 0);
  SetLength(arWordTextual, 0);
  iWordsUsed := 0;
  iWordsTotal := 0;
  stBookmark := '*';
  sbStatusBar.SimpleText := 'Concordance in processing. Total words: ' +
    IntToStr(iWordsTotal) + ' - Unique words: ' + IntToStr(iWordsUsed);
  sgWordList.RowCount := 1;
  lsContext.Clear;
  stText := stStartText + ' ';
  stText := StringReplace(stText, #9, ' ', [rfReplaceAll]);
  stText := StringReplace(stText, #10, ' ', [rfReplaceAll]);
  stText := StringReplace(stText, #13, ' ', [rfReplaceAll]);
  stText := StringReplace(stText, #39, #39 + ' ', [rfReplaceAll]);
  stText := StringReplace(stText, #96, #39 + ' ', [rfReplaceAll]);
  stText := StringReplace(stText, '  ', ' ', [rfReplaceAll]);
  try
    Screen.Cursor := crHourGlass;
    Application.ProcessMessages;
    slNoWord := TStringList.Create;
    slNoWord.CaseSensitive := False;
    slNoWord.CommaText := meSkipList.Text;
    while UTF8Pos(' ', stText) > 0 do
    begin
      stWord := UTF8Copy(stText, 1, UTF8Pos(' ', stText) - 1);
      if UTF8Pos('[[', stWord) > 0 then
      begin
        stBookmark := UTF8Copy(stWord, UTF8Pos('[[', stWord) + 2,
          UTF8Pos(']]', stWord) - UTF8Pos('[[', stWord) - 2);
        stText := UTF8Copy(stText, UTF8Pos(']]', stText) + 2, UTF8Length(stText));
        if UTF8Length(stBookmark) > 20 then
          stBookmark := UTF8Copy(stBookmark, 1, 20) + '...';
        Application.ProcessMessages;
        Continue;
      end;
      Inc(iWordsTotal);
      SetLength(arWordTextual, Length(arWordTextual) + 1);
      arWordTextual[Length(arWordTextual) - 1].stRecWordTextual := stWord;
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
            sbStatusBar.SimpleText :=
              'Concordance in processing. Total words: ' + IntToStr(iWordsTotal) +
              ' - Unique words: ' + IntToStr(iWordsUsed);
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
      stText := UTF8Copy(stText, UTF8Pos(' ', stText) + 1, UTF8Length(stText));
      Application.ProcessMessages;
    end;
  finally
    slNoWord.Free;
    sgWordList.Enabled := True;
    Screen.Cursor := crDefault;
  end;
  CompileGrid;
  ttEnd := Now;
  if sgWordList.RowCount > 1 then
  begin
    sbStatusBar.SimpleText :=
      'Total words: ' + IntToStr(iWordsTotal) + ' - Unique words: ' +
      IntToStr(iWordsUsed) + ' - Concordance done in ' +
      FormatDateTime('hh:nn:ss', ttEnd - ttStart) + '.';
    sgWordList.Row := 1;
    lsContext.Clear;
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
    sbStatusBar.SimpleText := 'No active concordance.';
    MessageDlg('No concordance available.', mtWarning, [mbOK], 0);
  end;
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
          stContext := stContext + arWordTextual[i].stRecWordTextual + '   '
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
    if sgWordList.Cells[0, sgWordList.Row] <> '' then
    begin
      meSkipList.Text := CleanField(meSkipList.Text);
      try
        WordsList := TStringList.Create;
        WordsList.CommaText := meSkipList.Text;
        if WordsList.IndexOf(sgWordList.Cells[0, sgWordList.Row]) < 0 then
        begin
          meSkipList.Text := meSkipList.Text + ', ' +
            sgWordList.Cells[0, sgWordList.Row];
          meSkipList.Text := CleanField(meSkipList.Text);
          sgWordList.DeleteRow(sgWordList.Row);
        end;
      finally
        WordsList.Free;
      end;
    end;
  end;
end;

procedure TfmMain.CompileGrid;
var
  i: integer;
begin
  // Compile grid of words
  if Length(arWordList) > 0 then
    sgWordList.Enabled := False;
  try
    Screen.Cursor := crHourGlass;
    sbStatusBar.SimpleText := 'List of words in processing, please, wait.';
    Application.ProcessMessages;
    SortWordFreq(arWordList, rgSortBy.ItemIndex);
    sgWordList.RowCount := 1;
    sgWordList.RowCount := Length(arWordList) + 1;
    for i := 0 to Length(arWordList) - 1 do
    begin
      sgWordList.Cells[0, i + 1] := arWordList[i].stRecWord;
      sgWordList.Cells[1, i + 1] := IntToStr(arWordList[i].iRecFreq);
      sgWordList.Cells[2, i + 1] := '0';
      ;
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
    sbStatusBar.SimpleText :=
      'Total words: ' + IntToStr(iWordsTotal) + ' - Unique words: ' +
      IntToStr(iWordsUsed) + ' - Concordance done in ' +
      FormatDateTime('hh:nn:ss', ttEnd - ttStart) + '.';
  finally
    sgWordList.Enabled := True;
    Screen.Cursor := crDefault;
  end;
end;

function TfmMain.CheckFilter(stWord: string): boolean;
var
  slFltStart, slFltEnd: TStringList;
  i: integer;
begin
  // Check the filtr on words
  Result := False;
  edFltStart.Text := CleanField(edFltStart.Text);
  edFltEnd.Text := CleanField(edFltEnd.Text);
  if ((edFltStart.Text = '') and (edFltEnd.Text = '')) then
  begin
    Result := True;
  end
  else if edFltStart.Text <> '' then
    try
      slFltStart := TStringList.Create;
      slFltStart.CommaText := edFltStart.Text;
      for i := 0 to slFltStart.Count - 1 do
      begin
        if UTF8LowerCase(UTF8Copy(stWord, 1, UTF8Length(slFltStart[i]))) =
          UTF8LowerCase(slFltStart[i]) then
        begin
          Result := True;
          Break;
        end;
      end;
    finally
      slFltStart.Free;
    end
  else if edFltEnd.Text <> '' then
    try
      slFltEnd := TStringList.Create;
      slFltEnd.CommaText := edFltEnd.Text;
      for i := 0 to slFltEnd.Count - 1 do
      begin
        if UTF8LowerCase(UTF8Copy(stWord, UTF8Length(stWord) -
          UTF8Length(slFltEnd[i]) + 1, UTF8Length(stWord))) =
          UTF8LowerCase(slFltEnd[i]) then
        begin
          Result := True;
          Break;
        end;
      end;
    finally
      slFltEnd.Free;
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
            if UTF8CompareStrCollated(arWordList[iPos2].stRecWord,
              arWordList[iPos2 + 1].stRecWord) > 0 then
            begin
              rcTempRec := arWordList[iPos2];
              arWordList[iPos2] := arWordList[iPos2 + 1];
              arWordList[iPos2 + 1] := rcTempRec;
            end;
          end;
        end;
      end
      else
      begin
        if UTF8CompareStrCollated(arWordList[iPos2].stRecWord,
          arWordList[iPos2 + 1].stRecWord) > 0 then
        begin
          rcTempRec := arWordList[iPos2];
          arWordList[iPos2] := arWordList[iPos2 + 1];
          arWordList[iPos2 + 1] := rcTempRec;
        end;
      end;
    end;
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
    myText := StringReplace(myText, ',', ', ', [rfReplaceAll]);
    myText := StringReplace(myText, ' ,', ',', [rfReplaceAll]);
    while Pos('  ', myText) > 0 do
      myText := StringReplace(myText, '  ', ' ', [rfReplaceAll]);
    if myText[Length(myText)] = ' ' then
      myText := Copy(myText, 0, Length(myText) - 1);
    if myText[1] = ' ' then
      myText := Copy(myText, 2, Length(myText));
    // If something wrong happens later...
    Result := myText;
    // Remove duplicates
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
    finally
      WordsList.Free;
      WordCurr.Free;
    end;
    Result := StringReplace(myText, LineEnding, ', ', [rfReplaceAll]);
    Result := UTF8Copy(Result, 1, UTF8Length(Result) - 2);
  end
  else
    Result := '';
end;

procedure TfmMain.SaveConcordance(stFileName: string);
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

function TfmMain.CleanXML(stText: string): string;
var
  blTag: boolean;
  i: integer;
begin
  // Clean a text from XML/HTML tags
  blTag := False;
  Result := '';
  for i := 0 to UTF8Length(stText) do
  begin
    if UTF8Copy(stText, i, 1) = '<' then
    begin
      blTag := True;
    end
    else if UTF8Copy(stText, i, 1) = '>' then
    begin
      blTag := False;
    end
    else if blTag = False then
    begin
      Result := Result + UTF8Copy(stText, i, 1);
    end;
  end;
  while UTF8Copy(Result, 1, 1) = LineEnding do
  begin
    Result := UTF8Copy(Result, 2, UTF8Length(Result));
  end;
end;

procedure TfmMain.CreateStatistic;
var
  i, n, iWord, iList, iTot: integer;
  slBookmark: TStringList;
begin
  // Create statistic
  if lbBookmarks.Count = 0 then
  begin
    MessageDlg('No bookmarks are set in the text.',
      mtWarning, [mbOK], 0);
    Abort;
  end;
  if sgWordList.RowCount = 1 then
  begin
    MessageDlg('No concordance available.',
      mtWarning, [mbOK], 0);
    Abort;
  end;
  try
    slBookmark := TStringList.Create;
    sgStatistic.RowCount := 1;
    sgStatistic.FixedRows := 1;
    sgStatistic.ColCount := lbBookmarks.Count + 1;
    sgStatistic.FixedCols := 1;
    for i := 1 to sgStatistic.ColCount - 1 do
    begin
      sgStatistic.ColWidths[i] := 70;
    end;
    for i := 0 to lbBookmarks.Count - 1 do
    begin
      sgStatistic.Cells[i + 1, 0] := lbBookmarks.Items[i];
    end;
    for i := 1 to sgWordList.RowCount - 1 do
    begin
      if sgWordList.Cells[2, i] = '1' then
      begin
        sgStatistic.RowCount := sgStatistic.RowCount + 1;
        sgStatistic.Cells[0, sgStatistic.RowCount - 1] :=
          sgWordList.Cells[0, i];
        slBookmark.CommaText := sgWordList.Cells[4, i];
        for n := 1 to sgStatistic.ColCount - 1 do
        begin
          iWord := 0;
          for iList := 0 to slBookmark.Count - 1 do
          begin
            if UTF8LowerCase(slBookmark[iList]) =
              sgStatistic.Cells[n, 0] then
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
    slBookmark.Free;
  end;
end;

end.

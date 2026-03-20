unit uPdfUtil;

interface
{$WARNINGS ON}
{$HINTS ON}
{$WARN UNIT_PLATFORM OFF}
{$WARN SYMBOL_PLATFORM OFF}

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ComCtrls, Buttons, FileCtrl,ExtCtrls,
  Jpeg,
  Printers,
  {$IFDEF UNIGUI}
    UFileUtilWeb, ApiUtilWeb,
  {$ELSE}
    UFileUtil, ApiUtil,
  {$ENDIF}

  MyUtilities, QuickPDF,
  EnDiGrph, EnImgScr, EnMisc, EnProLib, EnEncrypt, EnTgaGr, EnPngGr,
  EnPcxGr, EnJpgGr, EnDcxGr, EnTifGr, EnBmpGr, EnIcoGr, EnWmfGr, EnTransf;

Type
  TPagePageProp = record
    Width  : Integer;
    Height : Integer;
  end;

function RenderFileToJpgList(Pdf_FileName: String;
                             FromPage : Integer;
                             ToPage   : Integer;
                             ToPath   : String;
                             FileList : TStringList;
                             DpiToUse : Integer;
                             Var ErrorCode : Integer): Boolean;

function RenderPDFToMultiTiff(Pdf_FileName: String;
                              Tmp: TTiffGraphic;
                              DpiToUse : Integer;
                              Var ErrorCode : Integer): Boolean;

function RenderFirstPageToDib(Pdf_FileName: String;
                              Tmp: TDibGraphic;
                              DpiToUse : Integer;
                              Var ErrorCode : Integer): Boolean;

function RenderSpecificPageToJPG(Pdf_FileName: String;
                                 PageNo : Integer;
                                 TmpJPG: TJpegGraphic;
                                 DpiToUse : Integer;
                                 Var ErrorCode : Integer): Boolean;

function RenderSpecificPageToDib(Pdf_FileName: String;
                                 PageNo : Integer;
                                 TmpDIB: TDibGraphic;
                                 DpiToUse : Integer;
                                 Var ErrorCode : Integer): Boolean;

function RenderFirstPageToTif(Pdf_FileName: String;
                              Tif_FileName: String;
                              DpiToUse : Integer;
                              Var ErrorCode : Integer): Boolean;

Function GetPdfPagesCount(PdfFileName : String) : Integer;

function GetPdfPageProp(Pdf_FileName: String;
                        PageNo : Integer): TPagePageProp;
function GetRenderPdfPageProp(Pdf_FileName: String;
                              PageNo : Integer): TPagePageProp;

implementation

function RenderFileToJpgList(Pdf_FileName: String;
                             FromPage : Integer;
                             ToPage   : Integer;
                             ToPath   : String;
                             FileList : TStringList;
                             DpiToUse : Integer;
                             Var ErrorCode : Integer): Boolean;
Var
  PDFLibrary   : TQuickPDF;
  UnlockResult : Integer;
  JpgFileName  : String;
  CurCursor    : TCursor;
  FHndl, Rslt, CurPage, PageRef, Options : Integer;
begin
  Options := 1; {Jpg}
  Result := False;
  ErrorCode := 0;
  FileList.Clear;

  CurCursor := Screen.Cursor;
  Try
    Screen.Cursor := crHourGlass;
    SysUtils.ForceDirectories(ToPath);

    PDFLibrary := TQuickPDF.Create;
    Try
      UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
      if UnlockResult <> 1 then
        Exit;

      FHndl := PDFLibrary.DAOpenFileReadOnly(Pdf_FileName, '');
      IF FHndl > 0 Then
      begin
        Try
          for CurPage := FromPage to ToPage do
          begin
            PageRef := PDFLibrary.DAFindPage(FHndl, CurPage);
            IF PageRef > 0 Then
            begin
              JpgFileName := RemoveBackSlashChar(ToPath) + '\TmpJpg' + IntToStr(CurPage) + '.jpg';
              Windows.DeleteFile(PWideChar(JpgFileName));
              Rslt := PDFLibrary.DARenderPageToFile(FHndl,
                                                    PageRef,
                                                    Options,
                                                    DpiToUse,
                                                    JpgFileName);
              If Rslt = 1 Then
              begin
                FileList.Add(JpgFileName);
              end;
            end;
          end;
          Result := (PDFLibrary.DAGetPageCount(FHndl) = FileList.Count);
        Finally
          PDFLibrary.DACloseFile(FHndl);
        End;
      end
      else
      begin
        ErrorCode := PDFLibrary.LastErrorCode;
        Result := False;
      end;
    Finally
      FreeAndNil(PDFLibrary);
    End;
  Finally
    Screen.Cursor := CurCursor;
  End;
end;

function RenderPDFToMultiTiff(Pdf_FileName: String;
                              Tmp: TTiffGraphic;
                              DpiToUse : Integer;
                              Var ErrorCode : Integer): Boolean;
Var
  PdfFileName  : WideString;
  TmpJpg       : TJPegGraphic;
  SaveTif      : TTiffGraphic;
  DstPath      : String;
  FileList     : TStringList;
  CurJpg       : Integer;
  Stream       : TMemoryStream;
begin
  Result := False;
  ErrorCode := 0;
  Tmp.Clear;
  PdfFileName := Trim(Pdf_FileName);
  if Not FileExists(PdfFileName) then
    Exit;

  DstPath := RemoveBackSlashChar(GetWinTempDir) + '\HanibaalTmp';
  SysUtils.ForceDirectories(DstPath);

  FileList := TStringList.Create;
  Try
    If RenderFileToJpgList(PdfFileName, 1, GetPdfPagesCount(PdfFileName), DstPath, FileList, DpiToUse, ErrorCode) Then
    begin
      TmpJpg := TJPegGraphic.Create;
      Try
        SaveTif := TTiffGraphic.Create;
        Stream  := TMemoryStream.Create;
        For CurJpg := 0 To FileList.Count - 1 Do
        begin
          TmpJpg.LoadFromFile(FileList.Strings[CurJpg]);
          SaveTif.Assign(TmpJpg);
          if CurJpg = 0 then
          begin
            SaveTif.SaveToStream(Stream)
          end
          else
          begin
            SaveTif.AppendToStream(Stream);
          end;
        end;
      Finally
        TmpJpg.Free;
      End;
      Stream.Position := 0;
      Tmp.LoadFromStream(Stream);
    end;
  Finally
    if Assigned(SaveTif) then
      SaveTif.Free;
    if Assigned(Stream) then
      Stream.Free;
    FileList.Free;
  End;
end;

function RenderFirstPageToDib(Pdf_FileName: String;
                              Tmp: TDibGraphic;
                              DpiToUse : Integer;
                              Var ErrorCode : Integer): Boolean;
Var
  PdfFileName  : WideString;
  PDFLibrary   : TQuickPDF;
  MemoryStream : TMemoryStream;
  FHndl        : Integer;
  PageRef      : Integer;
  UnlockResult : Integer;
  TmpJpg       : TJPegGraphic;
  CurCursor    : TCursor;
begin
  Result := False;
  ErrorCode := 0;
  Tmp.Clear;
  PdfFileName := Trim(Pdf_FileName);
  if Not FileExists(PdfFileName) then
    Exit;

  Try
    PDFLibrary := TQuickPDF.Create;
    Try
      UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
      IF UnlockResult <> 1 then
        Exit;

      CurCursor := Screen.Cursor;
      Try
        Screen.Cursor := crHourGlass;
        FHndl := PDFLibrary.DAOpenFileReadOnly(PdfFileName,'');
        IF FHndl > 0 Then
        begin
          PageRef := PDFLibrary.DAFindPage(FHndl,1); // get first page reference
          IF PageRef > 0 Then
          begin
            MemoryStream := TMemoryStream.Create;
            Try
              PDFLibrary.DARenderPageToStream(FHndl,
                                              PageRef,
                                              1, {JPEG}
                                              DpiToUse, {DPIxy}
                                              MemoryStream {FileName});
              MemoryStream.Seek(0, soFromBeginning);
              Try
                TmpJpg := TJPegGraphic.Create;
                TmpJpg.LoadFromStream(MemoryStream);
                Tmp.Assign(TmpJpg);
                Result := Not Tmp.IsEmpty; // success
              Finally
                TmpJpg.Free;
              End;
            Finally
              MemoryStream.Free;
            End;
          end;
        end;
      Finally
        Screen.Cursor := CurCursor;
      End;
    Finally
      PDFLibrary.Free;
    End;
  Except;
    ErrorCode := PDFLibrary.LastErrorCode;
    Result := False; // error on proc
  End;
end;

function RenderSpecificPageToJPG(Pdf_FileName: String;
                                 PageNo : Integer;
                                 TmpJPG: TJpegGraphic;
                                 DpiToUse : Integer;
                                 Var ErrorCode : Integer): Boolean;
Var
  PdfFileName  : WideString;
  PDFLibrary   : TQuickPDF;
  MemoryStream : TMemoryStream;
  FHndl        : Integer;
  PageRef      : Integer;
  UnlockResult : Integer;
  CurCursor    : TCursor;
begin
  Result := False;
  ErrorCode := 0;
  TmpJPG.Clear;
  PdfFileName := Trim(Pdf_FileName);
  if Not FileExists(PdfFileName) then
    Exit;

  Try
    PDFLibrary := TQuickPDF.Create;
    Try
      UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
      IF UnlockResult <> 1 then
        Exit;

      CurCursor := Screen.Cursor;
      Try
        Screen.Cursor := crHourGlass;
        FHndl := PDFLibrary.DAOpenFileReadOnly(PdfFileName,'');
        IF FHndl > 0 Then
        begin
          PageRef := PDFLibrary.DAFindPage(FHndl,PageNo);
          IF PageRef > 0 Then
          begin
            MemoryStream := TMemoryStream.Create;
            Try
              PDFLibrary.DARenderPageToStream(FHndl,
                                              PageRef,
                                              1, {JPEG}
                                              DpiToUse, {DPIxy}
                                              MemoryStream {FileName});
              MemoryStream.Seek(0, soFromBeginning);
              Try
                TmpJpg.LoadFromStream(MemoryStream);
                Result := Not TmpJPG.IsEmpty; // success
              Finally
              End;
            Finally
              MemoryStream.Free;
            End;
          end;
        end;
      Finally
        Screen.Cursor := CurCursor;
      End;
    Finally
      PDFLibrary.Free;
    End;
  Except;
    ErrorCode := PDFLibrary.LastErrorCode;
    Result := False; // error on proc
  End;
end;

function RenderSpecificPageToDib(Pdf_FileName: String;
                                 PageNo : Integer;
                                 TmpDIB: TDibGraphic;
                                 DpiToUse : Integer;
                                 Var ErrorCode : Integer): Boolean;
Var
  PdfFileName  : WideString;
  PDFLibrary   : TQuickPDF;
  MemoryStream : TMemoryStream;
  FHndl        : Integer;
  PageRef      : Integer;
  UnlockResult : Integer;
  TmpJpg       : TJPegGraphic;
  CurCursor    : TCursor;
begin
  Result := False;
  ErrorCode := 0;
  TmpDIB.Clear;
  PdfFileName := Trim(Pdf_FileName);
  if Not FileExists(PdfFileName) then
    Exit;

  Try
    PDFLibrary := TQuickPDF.Create;
    Try
      UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
      IF UnlockResult <> 1 then
        Exit;

      CurCursor := Screen.Cursor;
      Try
        Screen.Cursor := crHourGlass;
        FHndl := PDFLibrary.DAOpenFileReadOnly(PdfFileName,'');
        IF FHndl > 0 Then
        begin
          PageRef := PDFLibrary.DAFindPage(FHndl,PageNo);
          IF PageRef > 0 Then
          begin
            MemoryStream := TMemoryStream.Create;
            Try
              PDFLibrary.DARenderPageToStream(FHndl,
                                              PageRef,
                                              1, {JPEG}
                                              DpiToUse, {DPIxy}
                                              MemoryStream {FileName});
              MemoryStream.Seek(0, soFromBeginning);
              Try
                TmpJpg := TJPegGraphic.Create;
                TmpJpg.LoadFromStream(MemoryStream);
                TmpDIB.Assign(TmpJpg);
                Result := Not TmpDIB.IsEmpty; // success
              Finally
                TmpJpg.Free;
              End;
            Finally
              MemoryStream.Free;
            End;
          end;
        end;
      Finally
        Screen.Cursor := CurCursor;
      End;
    Finally
      PDFLibrary.Free;
    End;
  Except;
    ErrorCode := PDFLibrary.LastErrorCode;
    Result := False; // error on proc
  End;
end;

function RenderFirstPageToTif(Pdf_FileName: String;
                              Tif_FileName: String;
                              DpiToUse : Integer;
                              Var ErrorCode : Integer): Boolean;
Var
  PdfFileName  : String;
  PDFLibrary   : TQuickPDF;
  UnlockResult : Integer;
  FHndl, Rslt, Page, PageRef, Options: Integer;
  CurCursor    : TCursor;
  DstPath      : String;
begin
  Page := 1;
  Options := 7;
  Result := False;
  ErrorCode := 0;

  PdfFileName := Trim(Pdf_FileName);
  if Not FileExists(PdfFileName) then
    Exit;
  If Not IsFileTiff(Tif_FileName) Then
    Exit;
  If FileExists(Tif_FileName) Then
    Windows.DeleteFile(PWideChar(Tif_FileName));
  If FileExists(Tif_FileName) Then // dont have inuf rights
    Exit;

  DstPath := j_PathFileName(Tif_FileName);
  SysUtils.ForceDirectories(DstPath);

  If Not SysUtils.DirectoryExists(DstPath) Then // dont have inuf rights
    Exit;

  CurCursor := Screen.Cursor;
  Try
    Screen.Cursor := crHourGlass;

    PDFLibrary := TQuickPDF.Create;
    Try
     UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
     if UnlockResult <> 1 then
       Exit;

      FHndl := PDFLibrary.DAOpenFileReadOnly(Pdf_FileName, '');
      IF FHndl > 0 Then
      begin
        Try
          PageRef := PDFLibrary.DAFindPage(FHndl, Page);
          IF PageRef > 0 Then
          begin
            Rslt := PDFLibrary.DARenderPageToFile(FHndl,
                                                  PageRef,
                                                  Options,
                                                  DpiToUse,
                                                  Tif_FileName);
            Result := (Rslt = 1);
          end;
        Finally
          PDFLibrary.DACloseFile(FHndl);
        End;
      end
      else
      begin
        ErrorCode := PDFLibrary.LastErrorCode;
        Result := False;
      end;
    Finally
      FreeAndNil(PDFLibrary);
    End;
  Finally
    Screen.Cursor := CurCursor;
  End;
end;

Function GetPdfPagesCount(PdfFileName : String) : Integer;
var
  PDFLibrary   : TQuickPDF;
  PDFHandler   : Integer;
  UnlockResult : Integer;
begin
  Result := -1;

  Try
    PDFLibrary := TQuickPDF.Create;
    Try
      UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
      if UnlockResult <> 1 then
        Exit;
      Try
        PDFHandler := PDFLibrary.DAOpenFileReadOnly(PdfFileName, '');
        IF PDFHandler > 0 Then
        begin
          Result := PDFLibrary.DAGetPageCount(PDFHandler);
        end;
      Finally
        PDFLibrary.DACloseFile(PDFHandler)
      End;
    except;
      Result := -1;
    end;
  Finally
    FreeAndNil(PDFLibrary);
  End;
end;

function GetPdfPageProp(Pdf_FileName: String;
                        PageNo : Integer): TPagePageProp;
Var
  PDFLibrary   : TQuickPDF;
  UnlockResult : Integer;
  JpgFileName  : String;
  CurCursor    : TCursor;
  FHndl, Rslt, CurPage, PageRef, Options : Integer;
begin
  Result.Width  := -1;
  Result.Height := -1;

  PDFLibrary := TQuickPDF.Create;
  Try
    UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
    if UnlockResult <> 1 then
      Exit;

    FHndl := PDFLibrary.LoadFromFile(Pdf_FileName, '');
    IF FHndl > 0 Then
    begin
      Try
       PDFLibrary.SelectPage(PageNo);
       Result.Width  := Trunc(PDFLibrary.PageWidth);
       Result.Height := Trunc(PDFLibrary.PageHeight);
      Finally
        PDFLibrary.DACloseFile(FHndl);
      End;
    end;
  Finally
    FreeAndNil(PDFLibrary);
  End;
end;

function GetRenderPdfPageProp(Pdf_FileName: String;
                              PageNo : Integer): TPagePageProp;
Var
  TmpJPG    : TJpegGraphic;
  ErrorCode : Integer;
begin
  Result.Width  := -1;
  Result.Height := -1;

  TmpJPG := TJpegGraphic.Create;
  Try
    if RenderSpecificPageToJPG(Pdf_FileName,
                               PageNo,
                               TmpJPG,
                               200, {DpiToUse}
                               ErrorCode) Then
    begin
      Result.Width  := TmpJPG.Width;
      Result.Height := TmpJPG.Height;
    end;
  Finally
    FreeAndNil(TmpJPG);
  End;

end;

end.

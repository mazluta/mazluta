unit uPdfUtil;

// =============================================================================
// uPdfUtil.pas - Converted from QuickPDF + Envision to ImageEN 13.x
// Original: QuickPDF Ver 17.1 + Envision Image Library 4.00
// Target  : ImageEN Ver 13.1.0 (Delphi 13)
// =============================================================================

interface
{$WARNINGS ON}
{$HINTS ON}
{$WARN UNIT_PLATFORM OFF}
{$WARN SYMBOL_PLATFORM OFF}

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ComCtrls, Buttons, FileCtrl, ExtCtrls,
  Jpeg,
  Printers,
  {$IFDEF UNIGUI}
    UFileUtilWeb, ApiUtilWeb,
  {$ELSE}
    UFileUtil, ApiUtil,
  {$ENDIF}
  MyUtilities,
  // ImageEN units (replaces QuickPDF + Envision)
  imageenio,        // TImageEnIO  - single image I/O
  imageenmio,       // TImageEnMIO - multi-image / multi-page I/O
  imageenproc,      // TImageEnProc - image processing
  iebitmap,         // TIEBitmap   - replaces TDibGraphic / TJpegGraphic / etc.
  iemultibitmap,    // TIEMultiBitmap - replaces TTiffGraphic (multi-page)
  hyiedefs,         // TIEPixelFormat, TIFF/JPEG compression constants
  hyiefunctions;    // Utility functions

// -----------------------------------------------------------------------------
// Type replaces:
//   TDibGraphic   -> TIEBitmap       (single image, any format)
//   TTiffGraphic  -> TIEMultiBitmap  (multi-page TIFF / PDF pages)
//   TJpegGraphic  -> TIEBitmap
// -----------------------------------------------------------------------------

Type
  TPagePageProp = record
    Width  : Integer;
    Height : Integer;
  end;

// Renders all PDF pages to individual JPG files in ToPath
function RenderFileToJpgList(Pdf_FileName: String;
                             FromPage : Integer;
                             ToPage   : Integer;
                             ToPath   : String;
                             FileList : TStringList;
                             DpiToUse : Integer;
                             Var ErrorCode : Integer): Boolean;

// Renders all PDF pages into a multi-bitmap (replaces TTiffGraphic)
function RenderPDFToMultiTiff(Pdf_FileName: String;
                              Tmp: TIEMultiBitmap;
                              DpiToUse : Integer;
                              Var ErrorCode : Integer): Boolean;

// Renders the first PDF page into a bitmap (replaces TDibGraphic)
function RenderFirstPageToDib(Pdf_FileName: String;
                              Tmp: TIEBitmap;
                              DpiToUse : Integer;
                              Var ErrorCode : Integer): Boolean;

// Renders a specific PDF page into a bitmap (replaces TJpegGraphic)
function RenderSpecificPageToJPG(Pdf_FileName: String;
                                 PageNo : Integer;
                                 TmpJPG: TIEBitmap;
                                 DpiToUse : Integer;
                                 Var ErrorCode : Integer): Boolean;

// Renders a specific PDF page into a bitmap (replaces TDibGraphic)
function RenderSpecificPageToDib(Pdf_FileName: String;
                                 PageNo : Integer;
                                 TmpDIB: TIEBitmap;
                                 DpiToUse : Integer;
                                 Var ErrorCode : Integer): Boolean;

// Renders the first PDF page and saves it as a TIFF file
function RenderFirstPageToTif(Pdf_FileName: String;
                              Tif_FileName: String;
                              DpiToUse : Integer;
                              Var ErrorCode : Integer): Boolean;

// Returns total page count of a PDF file (-1 on error)
Function GetPdfPagesCount(PdfFileName : String) : Integer;

// Returns the pixel dimensions of a rendered PDF page at 72 DPI
function GetPdfPageProp(Pdf_FileName: String;
                        PageNo : Integer): TPagePageProp;

// Returns the pixel dimensions of a PDF page rendered at 200 DPI
function GetRenderPdfPageProp(Pdf_FileName: String;
                              PageNo : Integer): TPagePageProp;

implementation

// =============================================================================
// Internal helper: render one PDF page into Bmp using ImageEN
// =============================================================================
function IE_RenderPdfPage(const PdfFileName : String;
                          PageNo            : Integer;
                          DpiToUse          : Integer;
                          Bmp               : TIEBitmap): Boolean;
var
  IEio : TImageEnIO;
begin
  Result := False;
  IEio   := TImageEnIO.Create(nil);
  try
    IEio.AttachIEBitmap(Bmp);
    IEio.Params.PDF_PageNum   := PageNo;   // 1-based page number
    IEio.Params.PDF_RenderDPI := DpiToUse;
    IEio.LoadFromFile(PdfFileName);
    Result := (Bmp.Width > 0) and (Bmp.Height > 0);
  finally
    IEio.Free;
  end;
end;

// =============================================================================
// Render all (or a range of) PDF pages to individual JPG files
// Replaces: QuickPDF DAOpenFileReadOnly + DAFindPage + DARenderPageToFile
// =============================================================================
function RenderFileToJpgList(Pdf_FileName: String;
                             FromPage : Integer;
                             ToPage   : Integer;
                             ToPath   : String;
                             FileList : TStringList;
                             DpiToUse : Integer;
                             Var ErrorCode : Integer): Boolean;
var
  TmpBmp      : TIEBitmap;
  IEio        : TImageEnIO;
  JpgFileName : String;
  CurCursor   : TCursor;
  CurPage     : Integer;
  PageCount   : Integer;
begin
  Result    := False;
  ErrorCode := 0;
  FileList.Clear;

  if not FileExists(Pdf_FileName) then
    Exit;

  PageCount := GetPdfPagesCount(Pdf_FileName);
  if PageCount <= 0 then
    Exit;

  CurCursor := Screen.Cursor;
  try
    Screen.Cursor := crHourGlass;
    SysUtils.ForceDirectories(ToPath);

    for CurPage := FromPage to Min(ToPage, PageCount) do
    begin
      JpgFileName := RemoveBackSlashChar(ToPath) + '\TmpJpg' + IntToStr(CurPage) + '.jpg';
      Windows.DeleteFile(PWideChar(JpgFileName));

      TmpBmp := TIEBitmap.Create;
      IEio   := TImageEnIO.Create(nil);
      try
        IEio.AttachIEBitmap(TmpBmp);
        IEio.Params.PDF_PageNum   := CurPage;
        IEio.Params.PDF_RenderDPI := DpiToUse;
        IEio.LoadFromFile(Pdf_FileName);
        if (TmpBmp.Width > 0) and (TmpBmp.Height > 0) then
        begin
          IEio.Params.JPEG_Quality := 85;
          IEio.SaveToFile(JpgFileName);
          if FileExists(JpgFileName) then
            FileList.Add(JpgFileName);
        end;
      finally
        IEio.Free;
        TmpBmp.Free;
      end;
    end;

    Result := (FileList.Count > 0);
  finally
    Screen.Cursor := CurCursor;
  end;
end;

// =============================================================================
// Render entire PDF into a TIEMultiBitmap (multi-page, replaces TTiffGraphic)
// Replaces: QuickPDF render + Envision TTiffGraphic.AppendToStream loop
// =============================================================================
function RenderPDFToMultiTiff(Pdf_FileName: String;
                              Tmp: TIEMultiBitmap;
                              DpiToUse : Integer;
                              Var ErrorCode : Integer): Boolean;
var
  IEMio     : TImageEnMIO;
  CurCursor : TCursor;
begin
  Result    := False;
  ErrorCode := 0;
  Tmp.Clear;

  if not FileExists(Pdf_FileName) then
    Exit;

  CurCursor := Screen.Cursor;
  try
    Screen.Cursor := crHourGlass;

    IEMio := TImageEnMIO.Create(nil);
    try
      IEMio.AttachIEMultiBitmap(Tmp);
      IEMio.Params.PDF_RenderDPI := DpiToUse;
      IEMio.LoadFromFile(Pdf_FileName);       // loads all PDF pages
      Result := (Tmp.ImageCount > 0);
    finally
      IEMio.Free;
    end;
  finally
    Screen.Cursor := CurCursor;
  end;
end;

// =============================================================================
// Render first PDF page into a TIEBitmap (replaces TDibGraphic)
// Replaces: QuickPDF DAOpenFileReadOnly + DARenderPageToStream + TJpegGraphic
// =============================================================================
function RenderFirstPageToDib(Pdf_FileName: String;
                              Tmp: TIEBitmap;
                              DpiToUse : Integer;
                              Var ErrorCode : Integer): Boolean;
var
  CurCursor : TCursor;
begin
  Result    := False;
  ErrorCode := 0;

  if not FileExists(Pdf_FileName) then
    Exit;

  CurCursor := Screen.Cursor;
  try
    Screen.Cursor := crHourGlass;
    Result := IE_RenderPdfPage(Pdf_FileName, 1, DpiToUse, Tmp);
  finally
    Screen.Cursor := CurCursor;
  end;
end;

// =============================================================================
// Render a specific PDF page into a TIEBitmap as JPEG
// Replaces: QuickPDF DARenderPageToStream + TJpegGraphic.LoadFromStream
// =============================================================================
function RenderSpecificPageToJPG(Pdf_FileName: String;
                                 PageNo : Integer;
                                 TmpJPG: TIEBitmap;
                                 DpiToUse : Integer;
                                 Var ErrorCode : Integer): Boolean;
var
  CurCursor : TCursor;
begin
  Result    := False;
  ErrorCode := 0;

  if not FileExists(Pdf_FileName) then
    Exit;

  CurCursor := Screen.Cursor;
  try
    Screen.Cursor := crHourGlass;
    Result := IE_RenderPdfPage(Pdf_FileName, PageNo, DpiToUse, TmpJPG);
  finally
    Screen.Cursor := CurCursor;
  end;
end;

// =============================================================================
// Render a specific PDF page into a TIEBitmap (DIB)
// Replaces: QuickPDF DARenderPageToStream + TDibGraphic.Assign(TJpegGraphic)
// =============================================================================
function RenderSpecificPageToDib(Pdf_FileName: String;
                                 PageNo : Integer;
                                 TmpDIB: TIEBitmap;
                                 DpiToUse : Integer;
                                 Var ErrorCode : Integer): Boolean;
var
  CurCursor : TCursor;
begin
  Result    := False;
  ErrorCode := 0;

  if not FileExists(Pdf_FileName) then
    Exit;

  CurCursor := Screen.Cursor;
  try
    Screen.Cursor := crHourGlass;
    Result := IE_RenderPdfPage(Pdf_FileName, PageNo, DpiToUse, TmpDIB);
  finally
    Screen.Cursor := CurCursor;
  end;
end;

// =============================================================================
// Render first PDF page and save directly as TIFF file
// Replaces: QuickPDF DARenderPageToFile with Options=7 (TIFF)
// =============================================================================
function RenderFirstPageToTif(Pdf_FileName: String;
                              Tif_FileName: String;
                              DpiToUse : Integer;
                              Var ErrorCode : Integer): Boolean;
var
  TmpBmp    : TIEBitmap;
  IEio      : TImageEnIO;
  CurCursor : TCursor;
  DstPath   : String;
begin
  Result    := False;
  ErrorCode := 0;

  if not FileExists(Pdf_FileName) then
    Exit;
  if not IsFileTiff(Tif_FileName) then   // IsFileTiff from ApiUtil / UFileUtil
    Exit;
  if FileExists(Tif_FileName) then
    Windows.DeleteFile(PWideChar(Tif_FileName));
  if FileExists(Tif_FileName) then        // delete failed - no rights
    Exit;

  DstPath := j_PathFileName(Tif_FileName);
  SysUtils.ForceDirectories(DstPath);
  if not SysUtils.DirectoryExists(DstPath) then
    Exit;

  CurCursor := Screen.Cursor;
  try
    Screen.Cursor := crHourGlass;

    TmpBmp := TIEBitmap.Create;
    IEio   := TImageEnIO.Create(nil);
    try
      IEio.AttachIEBitmap(TmpBmp);
      IEio.Params.PDF_PageNum   := 1;
      IEio.Params.PDF_RenderDPI := DpiToUse;
      IEio.LoadFromFile(Pdf_FileName);

      if (TmpBmp.Width > 0) and (TmpBmp.Height > 0) then
      begin
        IEio.Params.TIFF_Compression := ioTIFF_LZW;   // use CCITTFAX4 for BW
        IEio.SaveToFile(Tif_FileName);
        Result := FileExists(Tif_FileName);
      end;
    finally
      IEio.Free;
      TmpBmp.Free;
    end;
  finally
    Screen.Cursor := CurCursor;
  end;
end;

// =============================================================================
// Get total page count of a PDF file
// Replaces: QuickPDF DAOpenFileReadOnly + DAGetPageCount + DACloseFile
// NOTE: verify Params.PDF_PageCount property name in your ImageEN 13 install.
//       Alternatives: Params.PDF_Pages  or  Params.ImageCount
// =============================================================================
Function GetPdfPagesCount(PdfFileName : String) : Integer;
var
  IEio : TImageEnIO;
  Bmp  : TIEBitmap;
begin
  Result := -1;
  if not FileExists(PdfFileName) then
    Exit;
  try
    Bmp  := TIEBitmap.Create;
    IEio := TImageEnIO.Create(nil);
    try
      IEio.AttachIEBitmap(Bmp);
      // LoadFileInfo reads header only - no pixel rendering
      IEio.LoadFileInfo(PdfFileName);
      // TODO: verify exact property name for PDF page count in ImageEN 13:
      //   IEio.Params.PDF_PageCount   (most likely)
      //   IEio.Params.PDF_Pages
      //   IEio.Params.ImageCount
      Result := IEio.Params.PDF_PageCount;
    finally
      IEio.Free;
      Bmp.Free;
    end;
  except
    Result := -1;
  end;
end;

// =============================================================================
// Get pixel dimensions of a PDF page (rendered at 72 DPI)
// Replaces: QuickPDF LoadFromFile + SelectPage + PageWidth/PageHeight
// NOTE: returns pixel dimensions, not PDF points. To convert to PDF points:
//       points = pixels * 72 / 72 (at 72 DPI, 1 pixel = 1 point)
// =============================================================================
function GetPdfPageProp(Pdf_FileName: String;
                        PageNo : Integer): TPagePageProp;
var
  TmpBmp : TIEBitmap;
  IEio   : TImageEnIO;
begin
  Result.Width  := -1;
  Result.Height := -1;

  if not FileExists(Pdf_FileName) then
    Exit;

  TmpBmp := TIEBitmap.Create;
  IEio   := TImageEnIO.Create(nil);
  try
    IEio.AttachIEBitmap(TmpBmp);
    IEio.Params.PDF_PageNum   := PageNo;
    IEio.Params.PDF_RenderDPI := 72;     // 72 DPI: 1 pixel = 1 PDF point
    IEio.LoadFromFile(Pdf_FileName);
    if (TmpBmp.Width > 0) and (TmpBmp.Height > 0) then
    begin
      Result.Width  := TmpBmp.Width;
      Result.Height := TmpBmp.Height;
    end;
  finally
    IEio.Free;
    TmpBmp.Free;
  end;
end;

// =============================================================================
// Get pixel dimensions of a PDF page rendered at 200 DPI
// (for accurate size measurement after rendering)
// =============================================================================
function GetRenderPdfPageProp(Pdf_FileName: String;
                              PageNo : Integer): TPagePageProp;
var
  TmpBmp    : TIEBitmap;
  ErrorCode : Integer;
begin
  Result.Width  := -1;
  Result.Height := -1;

  TmpBmp := TIEBitmap.Create;
  try
    if RenderSpecificPageToDib(Pdf_FileName, PageNo, TmpBmp, 200, ErrorCode) then
    begin
      Result.Width  := TmpBmp.Width;
      Result.Height := TmpBmp.Height;
    end;
  finally
    TmpBmp.Free;
  end;
end;

end.

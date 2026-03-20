unit uAddToPdf;

interface
{$WARNINGS ON}
{$HINTS ON}
{$WARN UNIT_PLATFORM OFF}
{$WARN SYMBOL_PLATFORM OFF}

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ComCtrls, Buttons, FileCtrl,ExtCtrls,
  Jpeg, Math,
  uGlobalTypes,
  {$IFDEF UNIGUI}
    UFileUtilWeb, ApiUtilWeb, MyUtilitiesWeb,
  {$ELSE}
    UFileUtil, ApiUtil, MyUtilities,
  {$ENDIF}
  uPdfUtil, Printers,
  EnDiGrph,
  EnImgScr,
  EnMisc,
  EnProLib,
  EnEncrypt,
  EnTgaGr,       { for TTgaGraphic }
  EnPngGr,       { for TPngGraphic }
  EnPcxGr,       { for TPcxGraphic }
  EnJpgGr,       { for TJpegGraphic }
  EnDcxGr,       { for TDcxGraphic }
  EnTifGr,       { for TTiffGraphic }
  EnBmpGr,       { for TBitmapGraphic }
  EnIcoGr,       { for TIconGraphic }
  EnWmfGr,       { for TMetaFileGraphic }
  EnTransf,      { For TRotateTransform}
  //GIFImage,      { Tchart Gif Image}
  QuickPDF{,
  PDFium}
  ;

//type
//  WinIsWow64 = function( Handle: THandle; var Iret: BOOL ): Windows.BOOL; stdcall;

Const
  QuickPdf_NoError             : Integer = 01;

  Err_FromDir_Empty            : Integer = 02;
  Err_FromDir_NotExists        : Integer = 03;
  Err_ToDir_Empty              : Integer = 04;
  Err_ToDir_NotExists          : Integer = 05;
  Err_FromFile_Empty           : Integer = 06;
  Err_FromFile_NotExists       : Integer = 07;
  Err_ToFile_Empty             : Integer = 08;
  Err_ToFile_NotLegal          : Integer = 09;
  Err_ToFile_Exist_CannotDel   : Integer = 10;
  Err_FormatType_NotLegal      : Integer = 11;

  Err_FromFile_ExtentionNotTif : Integer = 12;
  Err_ToFile_ExtentionNotTif   : Integer = 13;

  Err_FromFile_ExtentionNotJpg : Integer = 14;
  Err_ToFile_ExtentionNotJpg   : Integer = 15;

  Err_FailToSaveFile           : Integer = 16;
  Err_LoadDPLRRender           : Integer = 97;
  Err_LoadPDFium32             : Integer = 98;
  Err_UnExceptedError          : Integer = 99;

Const
  SavePDF_InSize_A4 : String = '4';
  SavePDF_InSize_A3 : String = '3';

Type
  EAppError   = class(Exception);

type
  TPDFDimensions = Record
     Height : Integer;
     Width  : Integer;
     Found  : Boolean;
  end;

function GetPdfLibErrorStr(ErrorCode : Integer) : String;

function ClearPDFDimensions : TPDFDimensions;
function GetPDFDimensions(SrcFileName : WideString; PageNo : Integer) : TPDFDimensions;
Function CompareTwoDimension(PDFDimensions1 : TPDFDimensions; PDFDimensions2 : TPDFDimensions) : Boolean;

function AddDibToPdfFile(TifDib      : TDibGraphic;
                         DstFileName : WideString;
                         FitToPage   : Boolean;
                         AfterPage   : Integer) : LongInt;

function AddMultiImgToPdfFile(SrcFileName : WideString;
                              DstFileName : WideString;
                              FitToPage   : Boolean;
                              AfterPage   : Integer) : LongInt;

function AddImgToPdfFile(SrcFileName  : WideString;
                         DstFileName  : WideString;
                         FitToPage    : Boolean;
                         AfterPage   : Integer) : LongInt; OverLoad;

function AddImgToPdfFile(TifImg       : TTiffGraphic;
                         DstFileName  : WideString;
                         FitToPage    : Boolean;
                         AfterPage   : Integer) : LongInt; OverLoad;

function AddHTMLToPdfFile(HTMLText    : WideString;
                          DstFileName : WideString;
                          FitToPage   : Boolean;
                          AfterPage   : Integer) : Boolean;

function DoQuickMergeFiles(FirstFileName   : WideString;
                           SecondFileName  : WideString;
                           OutputFileName  : WideString) : Boolean;{ OverLoad;}
//function DoQuickMergeFiles(FirstFileName   : WideString;
//                           SecondFileName  : WideString;
//                           OutputFileName  : WideString;
//                           FitToPage       : Boolean) : Boolean; OverLoad;

function DoMergeFiles(FileToAdd   : WideString;
                      ToFileName  : WideString;
                      DstFileName : WideString;
                      DpiToUse    : Integer;
                      FitToPage   : Boolean) : Boolean;

function DoMergeFilesList(FileLists   : TStringList;
                          DstFileName : WideString;
                          DpiToUse    : Integer;
                          FitToPage   : Boolean) : Boolean;

function DoRotatePdf(SrcFile : String;
                     DstFile : String;
                     Page    : Integer;
                     Angle   : Integer) : Boolean;

function DoNormalizePage(SrcFile : String;
                         DstFile : String;
                         Page    : Integer) : Boolean;

function DoExtractFilePages(SrcFile : String;
                            DstFile : String;
                            StartPage : Integer;
                            PagesCount : Integer) : Boolean;

implementation

function IsFileWMF(FileName : String) : Boolean;
Var
  ExtName       : String;
begin
  ExtName := J_ExtFileName(FileName);
  Result  := FALSE;
  IF UpperCase(ExtName) = UpperCase('WMF') Then
    Result := TRUE;
end;

function IsFileBMP(FileName : String) : Boolean;
Var
  ExtName       : String;
begin
  ExtName := J_ExtFileName(FileName);
  Result  := FALSE;
  IF UpperCase(ExtName) = UpperCase('BMP') Then
    Result := TRUE;
end;

function IsFileJpeg(FileName : String) : Boolean;
Var
  ExtName       : String;
begin
  ExtName := J_ExtFileName(FileName);
  Result  := FALSE;
  IF (UpperCase(ExtName) = UpperCase(JPG_EXTENTION_NAME[1])) OR
     (UpperCase(ExtName) = UpperCase(JPG_EXTENTION_NAME[2])) Then
    Result := TRUE;
end;

function IsFileTiff(FileName : String) : Boolean;
Var
  ExtName       : String;
begin
  ExtName := J_ExtFileName(FileName);
  Result := True;
  IF (UpperCase(ExtName) <> UpperCase(TIF_EXTENTION_NAME[1])) And
     (UpperCase(ExtName) <> UpperCase(TIF_EXTENTION_NAME[2])) Then
    Result := False;
end;

function IsFileDCX(FileName : String) : Boolean;
Var
  ExtName : String;
begin
  ExtName := J_ExtFileName(FileName);
  Result := (UpperCase(ExtName) = UpperCase('Dcx'));
end;

function IsFilePDF(FileName : String) : Boolean;
Var
  ExtName : String;
begin
  ExtName := J_ExtFileName(FileName);
  Result := (UpperCase(ExtName) = UpperCase('PDF'));
end;

function GetPdfLibErrorStr(ErrorCode : Integer) : String;
begin
  Result := 'Unknown error';
  case ErrorCode of
    101 : Result := 'The Strength parameter passed to the Encrypt function was invalid';
    102 : Result := 'The Permissions parameter passed to the Encrypt function was invalid. ' +
                     ' Use the EncodePermissions function to construct a value for this parameter';
    103 : Result := 'The Encrypt function was used on a document that was already encrypted';
    104 : Result := 'The Encrypt function failed for an unknown reason';
    201 : Result := 'The SetInformation function failed because the document is encrypted';
    202 : Result := 'The Key parameter passed to the SetInformation function was out of range';
    301 : Result := 'An invalid combination of barcode and option was sent to the DrawBarcode function';
    302 : Result := 'Non-numeric characters were sent to DrawBarcode using EAN-13';
    303 : Result := 'The EAN-13 barcode has an invalid checksum character';
    401 : Result := 'Could not open input file';
    402 : Result := 'Output file already exists and could not be deleted';
    403 : Result := 'Could not open output file';
    404 : Result := 'Invalid password';
    405 : Result := 'Document is not encrypted';
    406 : Result := 'Document is already encrypted';
    407 : Result := 'Invalid encryption strength';
    408 : Result := 'Invalid permissions';
    409 : Result := 'Invalid file structure, file is damaged';
    410 : Result := 'One of the input files is encrypted';
    411 : Result := 'File not found';
    412 : Result := 'Invalid page range list';
    501 : Result := 'The specified FileHandle was invalid';
    999 : Result := 'The function could not be used because the library is not unlockedend';
  end;
end;

function GetPdfMergeTmpPath : String;
begin
  Result := RemoveBackSlashChar(GetWinTempDir) + '\HanibaalTmp\PdfTmp\' + IntToStr(GetTickCount);
  SysUtils.ForceDirectories(Result);
end;

function GetImagePageCount(FileName: String): Integer;
Var
  ExtName : String;
  CurImg  : Integer;
  FHndl   : Integer;
  Rslt    : Integer;
  UnlockResult : Integer;
  PDFLibrary   : TQuickPDF;
  SrcStream    : TFileStream;
  SrcDcx       : TDcxGraphic;
  SrcTif       : TTiffGraphic;
Const
  IMAGE_CAN_GET_PAGE_COUNT : Array [1..15] of String = ('TIF',
                                                        'TIFF',
                                                        'DCX',
                                                        'JPG',
                                                        'JPEG',
                                                        'BMP',
                                                        'PNG',
                                                        'ICO',
                                                        'PCX',
                                                        'PCC',
                                                        'GIF',
                                                        'TGA',
                                                        'VST',
                                                        'AFI',
                                                        'PDF');
begin
  Result := 1;
  IF Not FileExists(FileName) Then
   Exit;

  ExtName := J_ExtFileName(FileName);
  ExtName := Trim(ExtName);
  For CurImg := Low(IMAGE_CAN_GET_PAGE_COUNT) To High(IMAGE_CAN_GET_PAGE_COUNT) Do
  begin
    IF (UpperCase(IMAGE_CAN_GET_PAGE_COUNT[CurImg]) = UpperCase(ExtName)) Then
    begin
      IF IsFilePDF(FileName) Then
      begin
        PDFLibrary := TQuickPDF.Create;
        Try
          UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
          if UnlockResult <> 1 then
           Exit;

          FHndl := PDFLibrary.DAOpenFileReadOnly(FileName, '');
          IF FHndl > 0 Then
          begin
            Rslt := PDFLibrary.DAGetPageCount(Rslt);
          end;

          IF FHndl > 0 Then
          begin
            PDFLibrary.DACloseFile(FHndl);
            Result := Rslt;
          end;
        Finally
          FreeAndNil(PDFLibrary);
        End;
        Break;
      end
      else IF IsFileTiff(FileName) Then
      begin
        SrcStream := TFileStream.Create(FileName, fmShareDenyNone);
        SrcStream.Position := 0;
        Try
          SrcTif := TTiffGraphic.Create;
          Try
            Try
              Result := SrcTif.GetImageCount(SrcStream);
            Except;
              Result := 1;
            End;
          Finally
            Try
              SrcTif.Free;
            Except;
            End;
          End;
        Finally
          Try
            SrcStream.Free;
          Except;
          End;
          Try
            SrcStream := nil;
          Except;
          End;
        End;
        Break;
      end
      else IF IsFileDCX(FileName) Then
      begin
        SrcStream := TFileStream.Create(FileName, fmOpenRead);
        SrcStream.Position := 0;
        Try
          SrcDcx := TDcxGraphic.Create;
          Try
            Result := SrcDcx.GetImageCount(SrcStream);
          Finally
            Try
              SrcDcx.Free;
            Except;
            End;
          End;
        Finally
          Try
            SrcStream.Free;
          Except;
          End;

          Try
            SrcStream := nil;
          Except;
          End;
        End;
        Break;
      end
      else
      begin
        Result := 1;
      end;
    end;
  end;
end;

function GetTifFrameCount(FileName: String): Integer;
Var
  SrcStream: TFileStream;
  SrcTif: TTiffGraphic;
begin
  Result := 0;
  SrcStream := TFileStream.Create(FileName, fmOpenRead);
  SrcStream.Position := 0;
  Try
    SrcTif := TTiffGraphic.Create;
    Try
      Result := SrcTif.GetImageCount(SrcStream);
    Finally
      Try SrcTif.Free; Except; End;
    End;
  Finally
    Try SrcStream.Free; Except; End;
    Try SrcStream := nil; Except; End;
  End;
end;

function RenderPdfPageToStream(FileName : WideString; PageNo : Integer; DPI : Integer; MS : TMemoryStream) : Boolean;
Var
  PDFLibrary   : TQuickPDF;
  FHndl        : Integer;
  PageRef      : Integer;
  UnlockResult : Integer;
  Rslt         : Integer;
  //PathName     : WideString;
  //TmpFileName  : WideString;
  CurCursor    : TCursor;
begin
  Result := False;
  IF Not FileExists(FileName) Then
    Exit;
  if DPI < 72 then
    DPI := 72;
  if DPI > 300 then
    DPI := 300;

  CurCursor := Screen.Cursor;
  Try
    Screen.Cursor := crHourGlass;

      MS.Clear;
      PDFLibrary  := TQuickPDF.Create;
      Try
        UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
        IF UnlockResult <> 1 then
          Exit;

        Try
          FHndl := PDFLibrary.DAOpenFileReadOnly(FileName, '');
          IF FHndl > 0 Then
          begin
            PageRef := PDFLibrary.DAFindPage(FHndl, PageNo);
            IF PageRef > 0 Then
            begin
              Rslt := PDFLibrary.DARenderPageToStream(FHndl,
                                                      PageRef,
                                                      1,
                                                      DPI,
                                                      Ms);
              Result := (Rslt = 1);
            end;
          end;
        Finally
          PDFLibrary.DACloseFile(FHndl);
        End;
      Finally
        FreeAndNil(PDFLibrary);
      End;

      MS.Seek(0, soFromBeginning);
  Finally
    Screen.Cursor := CurCursor;
  End;
end;

function LoadImageToDib(FileName : String;
                        Var Tmp  : TDibGraphic;
                        PageNo   : Integer;
                        SilentMode: Boolean = False): Boolean;
Var
  MS      : TMemoryStream;
  TmpBMP  : TBitmapGraphic;
  TmpICO  : TIconGraphic;
  TmpJPG  : TJpegGraphic;
  TmpWmf  : TMetaFileGraphic;
  TmpPCX  : TPcxGraphic;
  TmpPNG  : TPngGraphic;
  TmpTGA  : TTgaGraphic;
  TmpTif  : TTiffGraphic;
  //TmpDib  : TDibGraphic;
  TmpDCX  : TDcxGraphic;
  TmpGif  : TImage;

  SrcRect : TRect;
  DstRect : TRect;

  ExtFileName : String;
  //SomeIcon    : TIcon;
  //TmpFileName : String;
begin
 // TBitmapGraphic   >>>>>> BMP
 // TIconGraphic     >>>>>> ICO
 // TJpegGraphic     >>>>>> JPG JPEG
 // TPcxGraphic      >>>>>> PCX PCC
 // TPngGraphic      >>>>>> PNG
 // TTgaGraphic      >>>>>> TGA VST AFI
 // TTiffGraphic     >>>>>> TIF TIFF

  Result := False;
  Tmp.Clear;

  IF Trim(FileName) = '' Then
  begin
    IF SilentMode Then
    begin
      Result := False;
      Exit;
    end
    else
      Raise EAppError.Create('קובץ לא נמצא - ' +  FileName);
  end;

  IF Not FileExists(FileName) Then
  begin
    IF SilentMode Then
    begin
      Result := true;
      Exit;
    end
    else
      Raise EAppError.Create('קובץ לא נמצא - ' + FileName);
  end;

  ExtFileName := ExtractFileExt(FileName);
  IF (copy(ExtFileName, 1, 1) = '.') Then
    ExtFileName := copy(ExtFileName, 2, Length(ExtFileName) - 1);

  IF Not FileExists(FileName) Then
   Exit;

  Try
    If UpperCase(ExtFileName) <> UpperCase('PDF') Then
    begin
      Try
        MS := TMemoryStream.Create;
        RenderPdfPageToStream(FileName, PageNo, 200, MS);
        MS.Position := 0;
        Tmp.LoadFromStream(MS);
      Finally
        FreeAndNil(MS);
      End;
    end
    else
    If (UpperCase(ExtFileName) = UpperCase('TIF')) Or
       (UpperCase(ExtFileName) = UpperCase('TIFF')) Then
    begin
      TmpTif := TTiffGraphic.Create;
      TmpTif.MultiLoad := False;
      TmpTif.ImageToLoad := PageNo;
      Try
        TmpTif.LoadFromFile(FileName);
        Tmp.Assign(TmpTif);
      Finally
        FreeAndNil(TmpTif);
      End;
    end
    else
    If (UpperCase(ExtFileName) = UpperCase('JPG')) Or
       (UpperCase(ExtFileName) = UpperCase('JPEG')) Then
    begin
      TmpJPG := TJpegGraphic.Create;
      Try
        TmpJPG.LoadFromFile(FileName);
        Tmp.Assign(TmpJPG);
      Finally
        FreeAndNil(TmpJPG);
      End;
    end
    else
    If UpperCase(ExtFileName) = UpperCase('PNG') Then
    begin
      TmpPNG := TPngGraphic.Create;
      Try
        TmpPNG.LoadFromFile(FileName);
        Tmp.Assign(TmpPNG);
      Finally
        FreeAndNil(TmpPNG);
      End;
    end
    else
    If UpperCase(ExtFileName) = UpperCase('BMP') Then
    begin
      TmpBMP := TBitmapGraphic.Create;
      Try
        TmpBMP.LoadFromFile(FileName);
        Tmp.Assign(TmpBMP);
      Finally
        FreeAndNil(TmpBMP);
      End;
    end
    else
    If UpperCase(ExtFileName) = UpperCase('DCX') Then
    begin
      TmpDCX := TDcxGraphic.Create;
      Try
        TmpDCX.LoadFromFile(FileName);
        Tmp.Assign(TmpDCX);
      Finally
        FreeAndNil(TmpDCX);
      End;
    end
    else
    If UpperCase(ExtFileName) = UpperCase('ICO') Then
    begin
      TmpICO := TIconGraphic.Create;
      Try
        TmpICO.LoadFromFile(FileName);
        Tmp.Assign(TmpICO);
      Finally
        FreeAndNil(TmpICO);
      End;
    end
    else
    If ((UpperCase(ExtFileName) = UpperCase('Wmf')) Or
        (UpperCase(ExtFileName) = UpperCase('Emf'))) Then
    begin
      TmpWmf := TMetaFileGraphic.Create;
      Try
        TmpWmf.LoadFromFile(FileName);
        Tmp.Assign(TmpWmf);
      Finally
        FreeAndNil(TmpWmf);
      End;
    end
    else
    If (UpperCase(ExtFileName) = UpperCase('PCX')) Or
       (UpperCase(ExtFileName) = UpperCase('PCC')) Then
    begin
      TmpPCX := TPcxGraphic.Create;
      Try
        TmpPCX.LoadFromFile(FileName);
        Tmp.Assign(TmpPCX);
      Finally
        FreeAndNil(TmpPCX);
      End;
    end
    else
    If (UpperCase(ExtFileName) = UpperCase('TGA')) Or
       (UpperCase(ExtFileName) = UpperCase('VST')) OR
       (UpperCase(ExtFileName) = UpperCase('AFI')) Then
    begin
      TmpTGA := TTgaGraphic.Create;
      Try
        TmpTGA.LoadFromFile(FileName);
        Tmp.Assign(TmpTGA);
      Finally
        FreeAndNil(TmpTGA);
      End;
    end
    else
    If (UpperCase(ExtFileName) = UpperCase('GIF')) Then
    begin
      Try
        TmpGif := TImage.Create(nil);
        TmpJPG := TJpegGraphic.Create;
        Try
          TmpGif.Picture.LoadFromFile(FileName);

          TmpJPG.NewImage(TmpGif.Width, TmpGif.Height, ifTrueColor, nil, 0, 0);
          SetStretchBltMode(TmpJPG.Canvas.Handle, STRETCH_DELETESCANS);

          SrcRect.Top := 0;
          SrcRect.Left := 0;
          SrcRect.Right := TmpGif.Width;
          SrcRect.Bottom := TmpGif.Height;

          DstRect.Top := 0;
          DstRect.Left := 0;
          DstRect.Right := TmpGif.Width;
          DstRect.Bottom := TmpGif.Height;

          TmpJPG.Canvas.CopyRect(DstRect, TmpGif.Canvas, SrcRect);

          Tmp.Assign(TmpJPG);
        Finally
          FreeAndNil(TmpGif);
          FreeAndNil(TmpJPG);
        End;
      Except;
        Tmp.Clear;
      End;
    end;

    Result := Not Tmp.IsEmpty;
  Except;
   Tmp.Clear;
   Result := False;
  End;
end;

function LoadFirstImageToTif(Filename : String;
                             Var Tmp : TTiffGraphic;
                             SilentMode : Boolean = FALSE) : Boolean;
Var
  TmpBMP      : TBitmapGraphic;
  TmpICO      : TIconGraphic;
  TmpJPG      : TJpegGraphic;
  TmpPCX      : TPcxGraphic;
  TmpPNG      : TPngGraphic;
  TmpTGA      : TTgaGraphic;
  TmpDib      : TDibGraphic;
  TmpGif      : TImage;
  ExtFileName : String;
  //FileLists   : TStringList;
  //DstPath     : String;
  //CurFile     : Integer;
  //JpgFileName : String;
  //Rslt        : Integer;
  //PDFLibrary  : TQuickPDF;
begin
  Result := FALSE;
  Tmp.ClearAll;

  //  TBitmapGraphic   >>>>>> BMP
  //  TIconGraphic     >>>>>> ICO
  //  TJpegGraphic     >>>>>> JPG JPEG
  //  TPcxGraphic      >>>>>> PCX PCC
  //  TPngGraphic      >>>>>> PNG
  //  TTgaGraphic      >>>>>> TGA VST AFI
  //  TTiffGraphic     >>>>>> TIF TIFF

  IF Not FileExists(FileName) Then
    Raise Exception.Create('File - ' + Filename + ' Not Found');

  Tmp.MultiLoad := FALSE;

  ExtFileName := J_ExtFileName(FileName);

  Try
    If UpperCase(ExtFileName) = UpperCase('PDF') Then
    begin
      Try
        TmpDib := TDibGraphic.Create;
        LoadImageToDib(Filename, TmpDib, 1);
        Tmp.Assign(TmpDib);
        Result := Not Tmp.IsEmpty;
      Finally
        FreeAndNil(TmpDib);
      End;
    end
    else
    If (UpperCase(ExtFileName) = UpperCase('TIF')) Or
       (UpperCase(ExtFileName) = UpperCase('TIFF')) Then
    begin
      Tmp.MultiLoad := False;
      Try
        Tmp.LoadFromFile(FileName);
      Except;
        TmpJPG := TJpegGraphic.Create;
        Try
          Try
            TmpJPG.LoadFromFile(FileName);
            Tmp.Assign(TmpJPG);
          Except;
            Tmp.ClearAll;
          End;
        Finally
          FreeAndNil(TmpJPG);
        End;
      End;
    end
    else
    If UpperCase(ExtFileName) = UpperCase('PNG') Then
    begin
      TmpPNG := TPngGraphic.Create;
      Try
        TmpPNG.LoadFromFile(FileName);
        Tmp.Assign(TmpPNG);
      Finally
        FreeAndNil(TmpPNG);
      End;
    end
    else
    If (UpperCase(ExtFileName) = UpperCase('JPG')) Or
       (UpperCase(ExtFileName) = UpperCase('JPEG')) Then
    begin
      TmpJPG := TJpegGraphic.Create;
      Try
        TmpJPG.LoadFromFile(FileName);
        Tmp.Assign(TmpJPG);
      Finally
        FreeAndNil(TmpJPG);
      End;
    end
    else
    If UpperCase(ExtFileName) = UpperCase('BMP') Then
    begin
      TmpBMP := TBitmapGraphic.Create;
      Try
        TmpBMP.LoadFromFile(FileName);
        Tmp.Assign(TmpBMP);
      Finally
        FreeAndNil(TmpBMP);
      End;
    end
    else
    If UpperCase(ExtFileName) = UpperCase('ICO') Then
    begin
      TmpICO := TIconGraphic.Create;
      Try
        TmpICO.LoadFromFile(FileName);
        Tmp.Assign(TmpICO);
      Finally
        FreeAndNil(TmpICO);
      End;
    end
    else
    If (UpperCase(ExtFileName) = UpperCase('PCX')) Or
       (UpperCase(ExtFileName) = UpperCase('PCC')) Then
    begin
      TmpPCX := TPcxGraphic.Create;
      Try
        TmpPCX.LoadFromFile(FileName);
        Tmp.Assign(TmpPCX);
      Finally
        FreeAndNil(TmpPCX);
      End;
    end
    else
    If (UpperCase(ExtFileName) = UpperCase('TGA')) Or
       (UpperCase(ExtFileName) = UpperCase('VST')) OR
       (UpperCase(ExtFileName) = UpperCase('AFI')) Then
    begin
      TmpTGA := TTgaGraphic.Create;
      Try
        TmpTGA.LoadFromFile(FileName);
        Tmp.Assign(TmpTGA);
      Finally
        FreeAndNil(TmpTGA);
      End;
    end
    else
    If (UpperCase(ExtFileName) = UpperCase('GIF')) Then
    begin
      TmpGif := TImage.Create(nil);
      Try
        TmpGif.Picture.LoadFromFile(FileName);
        Tmp.Assign(TmpGIF);
      Finally
        FreeAndNil(TmpGIF);
      End;
    end
    else
    begin
      TmpDib := TDibGraphic.Create;
      Try
        Try
          TmpDib.LoadFromFile(FileName);
          Tmp.Assign(TmpDib);
        Except;
          Tmp.ClearAll;
        End;
      Finally
        FreeAndNil(TmpDib);
      End;
    end;
    Result := Not Tmp.IsEmpty;
  Except;
    Result := False;
  End;
end;

procedure ConvertDibToGray(SrcDib : TDibGraphic; DstDib : TDibGraphic);
var
  Transform : TConvertToGrayTransform;
begin
  Transform := TConvertToGrayTransform.Create;
  try
    Transform.ApplyOnDest(SrcDib, DstDib);
  finally
    Transform.Free;
  end;
end;

function ClearPDFDimensions : TPDFDimensions;
begin
  with Result Do
  begin
    Height := -1;
    Width  := -1;
    Found  := False;
  end;
end;

function GetPDFDimensions(SrcFileName : WideString; PageNo : Integer) : TPDFDimensions;
Var
  TmpPDFLibrary : TQuickPDF;
  UnlockResult  : Integer;
  FHndl,
    //Rslt,
    //DPI,
    //CurPage,
    PageRef,
    fPageNo,
    PageCount :    Integer;
begin
  Result.Width  := -1;
  Result.Height := -1;
  Result.Found  := False;

  if not FileExists(SrcFileName) then
    Exit;

  PageCount := GetPdfPagesCount(SrcFileName);
  fPageNo   := PageNo;
  if (PageNo < 1) Or
     (PageNo > PageCount) Then
  begin
     fPageNo := 1;
  end;

  TmpPDFLibrary := TQuickPDF.Create;
  Try
    UnlockResult := TmpPDFLibrary.UnlockKey(edtLicenseKey);
    if UnlockResult <> 1 then
      Exit;

    FHndl := TmpPDFLibrary.DAOpenFileReadOnly(SrcFileName, '');
    IF FHndl > 0 Then
    begin
      TmpPDFLibrary.SetMeasurementUnits(0); //pixel
      PageRef := TmpPDFLibrary.DAFindPage(FHndl, fPageNo{PDFLibrary.DAGetPageCount(FHndl)});
      IF PageRef > 0 Then
      begin
        Result.Width  := Trunc(TmpPDFLibrary.DAGetPageWidth(FHndl,PageRef));
        Result.Height := Trunc(TmpPDFLibrary.DAGetPageHeight(FHndl,PageRef));
        Result.Found  := True;
      end;
    end;
    TmpPDFLibrary.DACloseFile(FHndl);
  Finally
    FreeAndNil(TmpPDFLibrary);
  End;
end;

Function CompareTwoDimension(PDFDimensions1 : TPDFDimensions; PDFDimensions2 : TPDFDimensions) : Boolean;
Var
  ChkDblWidth  : Double;
  ChkDblHeight : Double;
begin
  Result := False;

  If Not PDFDimensions1.Found Then
    Exit;
  If Not PDFDimensions2.Found Then
    Exit;

  If PDFDimensions1.Width < 10 Then
    Exit;
  If PDFDimensions2.Width < 10 Then
    Exit;
  If PDFDimensions1.Height < 10 Then
    Exit;
  If PDFDimensions2.Height < 10 Then
    Exit;

  ChkDblWidth  := PDFDimensions1.Width / PDFDimensions2.Width;
  ChkDblHeight := PDFDimensions1.Height / PDFDimensions2.Height;
  Result := ((ChkDblWidth > 0.98) And (ChkDblWidth < 1.02)) And
            ((ChkDblHeight > 0.98) And (ChkDblHeight < 1.02));
end;

function DoDrawImage(PDFLibrary   : TQuickPDF;
                     ImgIndex     : Integer;
                     fHeight      : Integer;
                     fWidth       : Integer;
                     fDstFileName : WideString) : Integer;
Var
  SavePath     : WideString;
begin
  Result := -1;
  Try
     PDFLibrary.SelectImage(ImgIndex);
     PDFLibrary.SetPageDimensions(fWidth,fHeight);
     Result := PDFLibrary.DrawImage(0, 0, fWidth,fHeight);
  Except;
    Result := -1;
  End;

  if Result = QuickPdf_NoError then
  begin
    Sleep(10);
    Application.ProcessMessages;
    SavePath := j_PathFileName(fDstFileName);
    if Not SysUtils.DirectoryExists(SavePath) then
      SysUtils.ForceDirectories(SavePath);

    Result := PDFLibrary.SaveToFile(fDstFileName);
    IF Result <> QuickPdf_NoError Then
      Result := Err_FailToSaveFile;
  end;
end;

function AddDibToPdfFile(TifDib      : TDibGraphic;
                         DstFileName : WideString;
                         FitToPage   : Boolean;
                         AfterPage   : Integer) : LongInt;
Var
  PDFLibrary    : TQuickPDF;
  PDFDimensions : TPDFDimensions;

  UnlockResult  : Integer;
  fHeight       : Integer;
  fWidth        : Integer;
  //Rslt          : Integer;

  ImgIndex      : Integer;
  TmpTiff       : TTiffGraphic;
  //TmpJpg        : TJpegGraphic;
  Stream        : TMemoryStream;
begin
  Result := Err_UnExceptedError;

  fHeight := TifDib.Height;
  fWidth  := TifDib.Width;

  IF FileExists(DstFileName) Then
  begin
    if FitToPage Then
    begin
      PDFDimensions := GetPDFDimensions(DstFileName, AfterPage);
      If PDFDimensions.Found then
      begin
        fWidth  := PDFDimensions.Width;
        fHeight := PDFDimensions.Height;
      end;
    end;
  end;

  PDFLibrary := TQuickPDF.Create;
  Try
    UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
    if UnlockResult <> 1 then
      Exit;

    IF Not FileExists(DstFileName) Then
    begin
      PDFLibrary.NewDocument;
    end
    else
    begin
      PDFLibrary.LoadFromFile(DstFileName,'');
      PDFLibrary.NewPage;
    end;

    PDFLibrary.SetOrigin(1);
    TmpTiff := TTiffGraphic.Create;
    Try
      TmpTiff.Assign(TifDib);
      Stream  := TMemoryStream.Create;
      Try
        TmpTiff.SaveToStream(Stream);
        ImgIndex := PDFLibrary.AddImageFromStream(Stream,0);
        Result := DoDrawImage(PDFLibrary,ImgIndex,fHeight,fWidth, DstFileName);
      Finally
        FreeAndNil(Stream);
      End;
    Finally
      FreeAndNil(TmpTiff);
    End;
  Finally
    FreeAndNil(PDFLibrary);
  End;
end;

function AddMultiImgToPdfFile(SrcFileName : WideString;
                              DstFileName : WideString;
                              FitToPage   : Boolean;
                              AfterPage   : Integer) : LongInt;
Var
  PDFLibrary     : TQuickPDF;
  UnlockResult   : Integer;
  IsFileExists   : Boolean;
  TifFrameCount  : Integer;
  fHeigh         : Integer;
  fWidth         : Integer;
  //Rslt          : Integer;
  PDFDimensions  : TPDFDimensions;
  NewFileName    : WideString;
  ListName       : WideString;
  OutputFileName : WideString;

  ImgIndex       : Integer;
  TmpTif         : TTiffGraphic;
  TmpDib         : TDibGraphic;
  //FrameNo       : Integer;
  CurFrame       : Integer;
  DstDir         : String;
begin
  Result := Err_UnExceptedError;
  If Not FileExists(SrcFileName) Then // the src file must exists
    Exit;
  If Not IsFilePDF(DstFileName) Then // the dest must be PDF file
    Exit;

  DstDir := j_PathFileName(DstFileName);
  if not DirectoryExists(DstDir) then
    ForceDirectories(DstDir);

  If IsFilePDF(SrcFileName) And
     IsFilePDF(DstFileName) And
     Not FileExists(DstFileName) then  // if dst not exists then copy src to dst
  begin
    if Not Windows.CopyFile(PWideChar(SrcFileName),PWideChar(DstFileName),False) then
      Raise Exception.Create('Fail Copy SrcFile: ' + SrcFileName + ' To ' + DstFileName);
    Result := QuickPdf_NoError;
    Exit; // Job Done - Dst not exists - so just copy Src to Dst
  end;

  If IsFilePDF(SrcFileName) And
     IsFilePDF(DstFileName) And
     Not FitToPage Then
  begin

    // if the files are PDF with no need to fit by size then Mearge Them and Go Out
    PDFLibrary := TQuickPDF.Create;
    Try
      UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
      if UnlockResult <> 1 then
        Exit;

      NewFileName := GetPdfMergeTmpPath + '\Tmp' + IntToStr(GetTickCount) + '.pdf';
      Result := PDFLibrary.ClearFileList(ListName);
      Result := PDFLibrary.AddToFileList(ListName,DstFileName); // this is the base file
      Result := PDFLibrary.AddToFileList(ListName,SrcFileName); // this is the file to add to the base
      Result := PDFLibrary.MergeFileListFast(ListName,NewFileName);

      If Result = 2 {The Number of files to be mearge... must be 2} Then
      begin
        if Not Windows.CopyFile(PWideChar(NewFileName),PWideChar(DstFileName),False) then
          Raise Exception.Create('Fail Mearge Files(1) ' + SrcFileName + ' and ' + DstFileName);
        Result := QuickPdf_NoError;
      end
      else
      begin
        Raise Exception.Create('Fail Mearge Files(2) ' + SrcFileName + ' and ' + DstFileName);
      end;
    Finally
      FreeAndNil(PDFLibrary);
    End;
    // leave the proc - work Done
    Exit;
  end; {Not FitToPage}

  // FitToPage = TRUE or Src is NOT PDF

  PDFLibrary := TQuickPDF.Create;
  Try
    UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
    if UnlockResult <> 1 then
      Exit;

    PDFDimensions := ClearPDFDimensions;
    IsFileExists  := FileExists(DstFileName);

    IF Not IsFileExists Then
    begin // New File
      PDFLibrary.NewDocument;
    end
    else
    begin  // File Exists - take old pdf param
      PDFLibrary.LoadFromFile(DstFileName,'');
      PDFDimensions := GetPDFDimensions(DstFileName,AfterPage);
    end;

    fHeigh := 30;
    fWidth := 30;
    PDFLibrary.SetOrigin(1);
    Try
      // get The Source Dimention
      // just to now the hight and the width of the image
      TmpTif := TTiffGraphic.Create;
      LoadFirstImageToTif(SrcFileName,
                          TmpTif,
                          True{SilentMode});
      fHeigh := TmpTif.Height;
      fWidth := TmpTif.Width;
    Finally
      TmpTif.Free;
    End;

    IF IsFileExists And
       FitToPage then
    begin
      if PDFDimensions.Found then
      begin // the Dst File Found - So we take the Dst PROP
        fHeigh := PDFDimensions.Height;
        fWidth := PDFDimensions.Width;
      end;
    end;

    IF IsFilePDF(SrcFileName) Then
    begin
      OutputFileName  := GetPdfMergeTmpPath + '\Tmp' + IntToStr(GetTickCount) + '.pdf';
      if DoQuickMergeFiles(DstFileName,
                           SrcFileName,
                           OutputFileName) Then
      begin
        Result := Err_FailToSaveFile;
        if Windows.DeleteFile(PWideChar(DstFileName)) Then
          if SysUtils.RenameFile(OutputFileName,DstFileName) then
            Result := QuickPdf_NoError;
      end
      else
      begin
        Result := Err_FailToSaveFile;
      end;
    end
    else
    IF IsFileTiff(SrcFileName) Then
    begin
      Try
        TmpTif := TTiffGraphic.Create;
        TmpDib := TDibGraphic.Create;

        TifFrameCount := GetTifFrameCount(SrcFileName);

        IF TifFrameCount = 0 Then
        begin
          TmpTif.LoadFromFile(SrcFileName);
          TmpDib.Assign(TmpTif);
          Result := AddDibToPdfFile(TmpDib,DstFileName,FitToPage,AfterPage);
        end
        else
        begin
          For CurFrame := 1 To TifFrameCount Do
          begin
            TmpTif.ClearAll;
            TmpTif.MultiLoad   := False;
            TmpTif.ImageToLoad := CurFrame;
            TmpTif.LoadFromFile(SrcFileName);
            TmpDib.Assign(TmpTif);
            Result := AddDibToPdfFile(TmpDib,DstFileName,FitToPage, AfterPage);
            if Result <> QuickPdf_NoError then
              Break;
          end; {For loop}
        end;
      Finally
        FreeAndNil(TmpDib);
        FreeAndNil(TmpTif);
      End;
    end
    else
    IF IsFilePNG(SrcFileName) Then
    begin
      IF IsFileExists Then
        PDFLibrary.NewPage;
      ImgIndex := PDFLibrary.AddImageFromFile(SrcFileName,0);
      Result := DoDrawImage(PDFLibrary,ImgIndex,fHeigh,fWidth,DstFileName);
    end
    else
    IF IsFileJpeg(SrcFileName) Then
    begin
      IF IsFileExists Then
        PDFLibrary.NewPage;
      ImgIndex := PDFLibrary.AddImageFromFile(SrcFileName,0);
      Result := DoDrawImage(PDFLibrary,ImgIndex,fHeigh,fWidth,DstFileName);
    end
    else
    IF IsFileBMP(SrcFileName) Then
    begin
      IF IsFileExists Then
        PDFLibrary.NewPage;
      ImgIndex := PDFLibrary.AddImageFromFile(SrcFileName,0);
      Result := DoDrawImage(PDFLibrary,ImgIndex,fHeigh,fWidth,DstFileName);
    end
    else
    IF IsFileWMF(SrcFileName) Then
    begin
      IF IsFileExists Then
        PDFLibrary.NewPage;
      ImgIndex := PDFLibrary.AddImageFromFile(SrcFileName,0);
      Result := DoDrawImage(PDFLibrary,ImgIndex,fHeigh,fWidth,DstFileName);
    end;
  Finally
    FreeAndNil(PDFLibrary);
  End;
end;

function AddImgToPdfFile(SrcFileName : WideString;
                         DstFileName : WideString;
                         FitToPage   : Boolean;
                         AfterPage   : Integer) : LongInt;
begin
  Result    := AddMultiImgToPdfFile(SrcFileName, DstFileName, FitToPage, AfterPage);
end;

function AddImgToPdfFile(TifImg      : TTiffGraphic;
                         DstFileName : WideString;
                         FitToPage   : Boolean;
                         AfterPage   : Integer) : LongInt;
Var
  TmpDib       : TDibGraphic;
  CurFrame     : Integer;
begin
  Result    := Err_UnExceptedError;

  TmpDib := TDibGraphic.Create;
  Try
    IF TifImg.FrameCount = 0 Then
    begin
      TmpDib.Assign(TifImg);
      Result := AddDibToPdfFile(TmpDib, DstFileName, FitToPage, AfterPage);
    end
    else
    begin
      For CurFrame := 1 To TifImg.FrameCount Do
      begin
        TmpDib.Assign(TifImg.Frames[CurFrame]);
        Result := AddDibToPdfFile(TmpDib, DstFileName, FitToPage, AfterPage);
        if Result <> QuickPdf_NoError then
          Break;
      end; {For loop}
    end;
  Finally
    FreeAndNil(TmpDib);
  End;
end;

function AddHTMLToPdfFile(HTMLText    : WideString;
                          DstFileName : WideString;
                          FitToPage   : Boolean;
                          AfterPage   : Integer) : Boolean;
Var
  PDFLibrary    : TQuickPDF;
  UnlockResult  : Integer;
  PDFDimensions : TPDFDimensions;
  fHeight       : Integer;
  fWidth        : Integer;
  //Rslt          : Integer;
begin
  Result := False;

  PDFLibrary := TQuickPDF.Create;
  Try
    UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
    if UnlockResult <> 1 then
      Exit;
    PDFLibrary.SetOrigin(1);

    fHeight       := 1000;
    fWidth        := 700;
    PDFDimensions := ClearPDFDimensions;
    IF Not FileExists(DstFileName) Then
    begin
      PDFLibrary.NewDocument;
    end
    else
    begin
      PDFDimensions := GetPDFDimensions(DstFileName,AfterPage);
      if FitToPage And PDFDimensions.Found then
      begin
        fWidth  := PDFDimensions.Width;
        fHeight := PDFDimensions.Height;
      end;
      PDFLibrary.LoadFromFile(DstFileName,'');
      PDFLibrary.NewPage;
    end;

    PDFLibrary.DrawHTMLText(0, 0, fWidth, HTMLText);
    PDFLibrary.SaveToFile(DstFileName);
  Finally
    FreeAndNil(PDFLibrary);
  End;
end;

function DoQuickMergeFiles(FirstFileName   : WideString;
                           SecondFileName  : WideString;
                           OutputFileName  : WideString) : Boolean;
Var
  DstPath      : String;
  PDFLibrary   : TQuickPDF;
  UnlockResult : Integer;
  Rslt         : Integer;
begin
  Result := False;

  If Not IsFilePDF(FirstFileName) Then
    Exit;
  If Not IsFilePDF(SecondFileName) Then
    Exit;
  If Not IsFilePDF(OutputFileName) Then
    Exit;

  if FileExists(OutputFileName) then
    Windows.DeleteFile(PWideChar(OutputFileName));
  if FileExists(OutputFileName) then
    Exit;

  DstPath := j_PathFileName(OutputFileName);
  if not SysUtils.DirectoryExists(DstPath) Then
    SysUtils.ForceDirectories(DstPath);
  if not SysUtils.DirectoryExists(DstPath) Then
    Exit;

  PDFLibrary := TQuickPDF.Create;
  Try
    UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
    if UnlockResult <> 1 then
      Exit;

    Rslt := PDFLibrary.MergeFiles(FirstFileName,SecondFileName,OutputFileName);
    Result := (Rslt = 1);
  Finally
    FreeAndNil(PDFLibrary);
  End;
end;

//function DoQuickMergeFiles(FirstFileName   : WideString;
//                           SecondFileName  : WideString;
//                           OutputFileName  : WideString;
//                           FitToPage       : Boolean) : Boolean;
//Var
//  PDFLibrary   : TQuickPDF;
//  UnlockResult : Integer;
//  Height       : Double;
//  Width        : Double;
//  NewHeight    : Double;
//  NewWidth     : Double;
//  CurPage      : Integer;
//  scaleX       : Double;
//  scaleY       : Double;
//  scale        : Double;
//  PDFDimensions : TPDFDimensions;
//begin
//  Result := True;
//  if not DoQuickMergeFiles(FirstFileName,
//                           SecondFileName,
//                           OutputFileName) Then
//  begin
//    Result := False;
//    Exit;
//  end;
//
//  if FitToPage then
//  begin
//    PDFLibrary := TQuickPDF.Create;
//    Try
//      UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
//      if UnlockResult <> 1 then
//        Exit;
//
//      PDFDimensions := GetPDFDimensions(OutputFileName, 1);
//      PDFLibrary.LoadFromFile(OutputFileName,'');
//
//      for CurPage := 1 to PDFLibrary.PageCount do
//      begin
//        PDFLibrary.SelectPage(CurPage);
//        if CurPage = 1 then
//        begin
//          Width  := PDFLibrary.GetPageBox(1, {MediaBox}
//                                          2 {Width});
//          Height := PDFLibrary.GetPageBox(1, {MediaBox}
//                                          3 {Height});
//
//          //if Width <> PDFDimensions.Width then
//          //  ShowMessage('check W');
//          //if Height <> PDFDimensions.Height then
//          //  ShowMessage('check H');
//
//        end
//        else
//        begin
//          NewWidth  := PDFLibrary.GetPageBox(1, {MediaBox}
//                                          2 {Width});
//          NewHeight := PDFLibrary.GetPageBox(1, {MediaBox}
//                                          3 {Height});
//          if ((NewWidth / Width) < 0.95) or
//             ((NewWidth / Width) > 1.05) or
//             ((NewHeight / Height) < 0.95) or
//             ((NewHeight / Height) > 1.05) then
//          begin
//
//            scaleX := Width / NewWidth;
//            scaleY := Height / NewHeight;
//
//            // For uniform scaling (keeps aspect ratio):
//            scale := Min(scaleX, scaleY);
//            PDFLibrary.SetRenderScale(scale);
//            PDFLibrary.SetPageDimensions(Width,Height)
//          end;
//        end;
//      end;
//      PDFLibrary.SaveToFile(OutputFileName);
//    Finally
//      FreeAndNil(PDFLibrary);
//    End;
//  end;
//end;

function DoMergeFiles(FileToAdd   : WideString;
                      ToFileName  : WideString;
                      DstFileName : WideString;
                      DpiToUse    : Integer;
                      FitToPage   : Boolean) : Boolean;
Var
  FileLists       : TStringList;
  CurFrame        : Integer;
  DstPath         : String;

  NewFileToAdd    : String;
  NewToFileName   : String;
  NewDstFileName  : String;
  NewDstDir       : String;

  JpgFileName     : String;
  PDFDimensions1  : TPDFDimensions;
  PDFDimensions2  : TPDFDimensions;
  ImgPages        : Integer;
  CurFile         : Integer;
  Rslt            : Integer;
  TmpPath         : String;
  TmpTif          : TTiffGraphic;
  TmpDib          : TDibGraphic;
  AfterPage       : Integer;
  ErrorCode       : Integer;
begin
  Result := False;
  AfterPage := -1;

  If Not IsFilePDF(DstFileName) Then
    Exit;

  if FileExists(DstFileName) then
    Windows.DeleteFile(PWideChar(DstFileName));
  if FileExists(DstFileName) then
    Exit;

  DstPath := j_PathFileName(DstFileName);
  if not SysUtils.DirectoryExists(DstPath) Then
    SysUtils.ForceDirectories(DstPath);
  if not SysUtils.DirectoryExists(DstPath) Then
    Exit;

  TmpPath         := RemoveBackSlashChar(GetWinTempDir) + '\HanibaalTmp\PdfTmp\' + IntToStr(GetTickCount);
  if not SysUtils.DirectoryExists(TmpPath) then
    SysUtils.ForceDirectories(TmpPath);
  if not SysUtils.DirectoryExists(TmpPath) Then
    Exit;

  NewFileToAdd   := RemoveBackSlashChar(TmpPath) + '\' + j_FileName(FileToAdd);
  NewToFileName  := RemoveBackSlashChar(TmpPath) + '\' + j_FileName(ToFileName);
  NewDstFileName := RemoveBackSlashChar(TmpPath) + '\' + j_FileName(DstFileName);
  Windows.DeleteFile(PWideChar(NewDstFileName));

  Windows.CopyFile(PWideChar(FileToAdd),PWideChar(NewFileToAdd), False);
  Windows.CopyFile(PWideChar(ToFileName),PWideChar(NewToFileName), False);

  FileLists := TStringList.Create;
  Try
    if Not FitToPage then
    begin
      FileLists.Clear;
      FileLists.Add(NewToFileName);
      FileLists.Add(NewFileToAdd);

      Result := DoMergeFilesList(FileLists, NewDstFileName, DpiToUse, FitToPage);
    end
    else
    begin
    //********************************************************************
    // FitToPage = TRUE - So, Src File Will Mearge to size of Dst File
    //********************************************************************
      if IsFilePDF(NewToFileName)  And   // if 2 files are PDF then merge PDF Files
         IsFilePDF(NewFileToAdd) then
      begin
        PDFDimensions1 := GetPDFDimensions(NewToFileName, AfterPage);
        PDFDimensions2 := GetPDFDimensions(NewFileToAdd, AfterPage);
        if CompareTwoDimension(PDFDimensions1, PDFDimensions2) Then
        begin
          FileLists.Clear;
          FileLists.Add(NewToFileName);
          FileLists.Add(NewFileToAdd);

          Result := DoMergeFilesList(FileLists, NewDstFileName, DpiToUse, FitToPage);
        end
        else
        begin
          If Not windows.CopyFile(PWideChar(NewToFileName), PWideChar(NewDstFileName), False) then
            Exit;
          if Not FileExists(NewDstFileName) then
            Exit;

          DstPath := GetPdfMergeTmpPath;
          If RenderFileToJpgList(NewFileToAdd,
                                 1,
                                 GetPdfPagesCount(NewFileToAdd),
                                 DstPath,
                                 FileLists,
                                 DpiToUse,
                                 ErrorCode) Then
          begin
            for CurFile := 0 To FileLists.Count - 1 Do
            begin
              JpgFileName := FileLists.Strings[CurFile];
              Rslt := AddImgToPdfFile(JpgFileName,
                                      NewDstFileName,
                                      FitToPage, {here - alwes true}
                                      AfterPage);
              if Rslt <> QuickPdf_NoError then
                Break;  // one of the files did not mearge
            end;

            Result := (Rslt = QuickPdf_NoError); // all images from Src file added to the Dst file
          end;
        end;
      end
      else
      begin
        if IsFilePDF(NewToFileName)  Then // if Dest files is PDF then the secend file is not
        begin
          If Not windows.CopyFile(PWideChar(NewToFileName), PWideChar(NewDstFileName), False) then
            Exit;
          if Not FileExists(NewDstFileName) then
            Exit;

          ImgPages := GetImagePageCount(NewFileToAdd);

          TmpTif := TTiffGraphic.Create;
          Try
            if ImgPages = 1 Then
            begin
              LoadFirstImageToTif(NewFileToAdd,TmpTif);
              Rslt := AddImgToPdfFile(TmpTif,NewDstFileName,FitToPage, AfterPage);
            end
            else
            begin
              Try
                TmpDib := TDibGraphic.Create;
                For CurFrame := 1 To ImgPages Do
                begin
                  TmpDib.Clear;
                  LoadImageToDib(NewFileToAdd, TmpDib, CurFrame);
                  TmpTif.ClearAll;
                  TmpTif.Assign(TmpDib);
                  Rslt := AddImgToPdfFile(TmpTif, NewDstFileName, FitToPage, AfterPage);
                  if Rslt <> QuickPdf_NoError then
                    Break;  // one of the files did not mearge
                end;
              Finally
                FreeAndNil(TmpDib);
              End;
            end;
            Result := (Rslt = QuickPdf_NoError);
          Finally
            TmpTif.Free;
          End;
        end
        else
        begin // the Dst file is NOT PDF - the Src can be and can be NOT PDF file
          TmpTif := TTiffGraphic.Create;
          Try
            ImgPages := GetImagePageCount(NewToFileName);
              // create the base PDF file From Dst File (wich is not PDF)
            if ImgPages = 1 Then
            begin
              LoadFirstImageToTif(NewToFileName,TmpTif);
              Rslt := AddImgToPdfFile(TmpTif, NewDstFileName, FitToPage, AfterPage);
            end
            else
            begin
              Try
                TmpDib := TDibGraphic.Create;
                For CurFrame := 1 To ImgPages Do
                begin
                  TmpDib.Clear;
                  LoadImageToDib(NewToFileName, TmpDib, CurFrame);
                  TmpTif.ClearAll;
                  TmpTif.Assign(TmpDib);
                  Rslt := AddImgToPdfFile(TmpTif, NewDstFileName, FitToPage, AfterPage);
                  if Rslt <> QuickPdf_NoError then
                    Break;  // one of the files did not mearge
                end;
              Finally
                FreeAndNil(TmpDib);
              End;

            end;

            Result := (Rslt = QuickPdf_NoError);
          Finally
            TmpTif.Free;
          End;

          If Not Result Then // we could not convert first file to PDF
            Exit;
          Result := False; // now First file is PDF (in NewDstFileName)

          // now we have the Dst As PDF
          // so, Add The Src To Dst
          if IsFilePDF(NewFileToAdd)  Then
          begin
            FileLists.Clear;
            DstPath := GetPdfMergeTmpPath;
            If RenderFileToJpgList(NewFileToAdd,
                                   1,
                                   GetPdfPagesCount(NewFileToAdd),
                                   DstPath,
                                   FileLists,
                                   DpiToUse,
                                   ErrorCode) Then
            begin
              for CurFile := 0 To FileLists.Count - 1 Do
              begin
                JpgFileName := FileLists.Strings[CurFile];
                Rslt := AddImgToPdfFile(JpgFileName,
                                        NewDstFileName,
                                        FitToPage {here - alwes true},
                                        AfterPage);
                if Rslt <> QuickPdf_NoError then
                  Break;  // one of the files did not mearge
              end;

              Result := (Rslt = QuickPdf_NoError); // all images from secend file added to the first file
            end;
          end
          else
          begin
            TmpTif := TTiffGraphic.Create;
            Try
              ImgPages := GetImagePageCount(NewFileToAdd);
                // create the base PDF file From Dst File (wich is not PDF)
              if ImgPages = 1 Then
              begin
                LoadFirstImageToTif(NewFileToAdd,TmpTif);
                Rslt := AddImgToPdfFile(TmpTif, NewDstFileName, FitToPage, AfterPage);
              end
              else
              begin
                Try
                  TmpDib := TDibGraphic.Create;
                  For CurFrame := 1 To ImgPages Do
                  begin
                    TmpDib.Clear;
                    LoadImageToDib(NewFileToAdd, TmpDib, CurFrame);
                    TmpTif.ClearAll;
                    TmpTif.Assign(TmpDib);
                    Rslt := AddImgToPdfFile(TmpTif, NewDstFileName, FitToPage, AfterPage);
                    if Rslt <> QuickPdf_NoError then
                      Break;  // one of the files did not mearge
                  end;
                Finally
                  FreeAndNil(TmpDib);
                End;
              end;
              Result := (Rslt = QuickPdf_NoError);
            Finally
              TmpTif.Free;
            End;
          end;
        end;
      end;
    end;

    if Result then // for now all good - copy newDst to Dst
    begin
      Result := False;
      if FileExists(NewDstFileName) then
      begin
        CopyFile(PWideChar(NewDstFileName),PWideChar(DstFileName), False);
        Result := FileExists(DstFileName);
      end;
    end;

  Finally
    FileLists.Free;
  End;
end;

function DoMergeFilesList(FileLists   : TStringList;
                          DstFileName : WideString;
                          DpiToUse    : Integer;
                          FitToPage   : Boolean) : Boolean;
Var
  PDFLibrary   : TQuickPDF;
  UnlockResult : Integer;
  Rslt         : Integer;
  ListName     : WideString;
  OldName      : WideString;
  NewName      : WideString;
  CurFile      : Integer;
  //WithColor    : Integer;
  SavePath     : String;
  AfterPage    : Integer;
begin
  Result    := False;
  AfterPage := -1;

  SavePath := GetPdfMergeTmpPath;

  PDFLibrary := TQuickPDF.Create;
  Try
    UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
    if UnlockResult <> 1 then
      Exit;

    if FileLists.Count > 1 then
    begin
      Rslt := PDFLibrary.ClearFileList(ListName);
      For CurFile := 0 To FileLists.Count - 1 Do
      begin
        OldName := FileLists.Strings[CurFile];
        if Not IsFilePdf(OldName) then
        begin
          NewName := RemoveBackSlashChar(SavePath) + '\Tmp' + IntToStr(GetTickCount) + '.pdf';
          AddMultiImgToPdfFile(OldName,
                               NewName,
                               FitToPage,
                               AfterPage);
        end
        else
          NewName := OldName;
        Rslt := PDFLibrary.AddToFileList(ListName,NewName);

        Rslt := PDFLibrary.MergeFileListFast(ListName,DstFileName);
        Result := (Rslt = FileLists.Count);
      end;
    end
    else
    begin
      FreeAndNil(PDFLibrary);
      OldName := FileLists.Strings[0]; // only one - so take the first
      if Not IsFilePdf(OldName) then
      begin
        AddMultiImgToPdfFile(OldName,
                             DstFileName,
                             FitToPage,
                             AfterPage);
      end
      else
      begin
        DoMergeFiles(OldName,DstFileName,DstFileName, DpiToUse, FitToPage);
      end;
    end;
  Finally
    FreeAndNil(PDFLibrary);
  End;
end;

function DoRotatePdf(SrcFile : String;
                     DstFile : String;
                     Page    : Integer;
                     Angle   : Integer) : Boolean;
Var
  PDFLibrary    : TQuickPDF;
  UnlockResult  : Integer;
  //PDFDimensions : TPDFDimensions;
  DstPath       : String;
  //Rslt          : Integer;
begin
  Result := False;
  if Trim(SrcFile) = '' Then
    exit;
  if Trim(DstFile) = '' Then
    exit;

  if not FileExists(Trim(SrcFile)) Then
    exit;
  if not IsFilePDF(SrcFile) then
    exit;
  if not IsFilePDF(DstFile) then
    exit;

  DstPath := j_PathFileName(DstFile);
  if not SysUtils.DirectoryExists(DstPath) then
    SysUtils.ForceDirectories(DstPath);
  if not SysUtils.DirectoryExists(DstPath) then
    exit;

  PDFLibrary := TQuickPDF.Create;
  Try
    UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
    if UnlockResult <> 1 then
      Exit;

    PDFLibrary.LoadFromFile(SrcFile,'');
    if PDFLibrary.SelectPage(Page) = 1 then
    begin
      if PDFLibrary.RotatePage(Angle) = 1 then
      begin
        if PDFLibrary.SaveToFile(DstFile) = 1 Then
          Result := True;
      end;
    end;
  Finally
    FreeAndNil(PDFLibrary);
  End;
end;

function DoNormalizePage(SrcFile : String;
                         DstFile : String;
                         Page    : Integer) : Boolean;
Var
  PDFLibrary    : TQuickPDF;
  UnlockResult  : Integer;
  //PDFDimensions : TPDFDimensions;
  DstPath       : String;
  //Rslt          : Integer;
  CurPage       : Integer;
  SwSave        : Boolean;
begin
  Result := False;
  if Trim(SrcFile) = '' Then
    exit;
  if Trim(DstFile) = '' Then
    exit;

  if not FileExists(Trim(SrcFile)) Then
    exit;
  if not IsFilePDF(SrcFile) then
    exit;
  if not IsFilePDF(DstFile) then
    exit;

  DstPath := j_PathFileName(DstFile);
  if not SysUtils.DirectoryExists(DstPath) then
    SysUtils.ForceDirectories(DstPath);
  if not SysUtils.DirectoryExists(DstPath) then
    exit;

  PDFLibrary := TQuickPDF.Create;
  Try
    UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
    if UnlockResult <> 1 then
      Exit;

    SwSave := True;
    PDFLibrary.LoadFromFile(SrcFile,'');

    if Page = ITEM_NOT_FOUND then
    begin
      for CurPage := 1 to PDFLibrary.PageCount do
      begin
        if PDFLibrary.SelectPage(CurPage) = 1 then
        begin
          If PDFLibrary.NormalizePage(3) <> 1 then
          begin
            SwSave := False;
            Break;
          end;
        end;
      end;
    end
    else
    begin
      if PDFLibrary.SelectPage(Page) = 1 then
      begin
        if PDFLibrary.NormalizePage(3) <> 1 then
        begin
          SwSave := False;
        end;
      end;
    end;

    if SwSave Then
    begin
      if PDFLibrary.SaveToFile(DstFile) = 1 Then
        Result := True;
    end
    else
    begin
      raise Exception.Create('Fail save PDF file');
    end;
  Finally
    FreeAndNil(PDFLibrary);
  End;
end;

function DoExtractFilePages(SrcFile : String;
                            DstFile : String;
                            StartPage : Integer;
                            PagesCount : Integer) : Boolean;
Var
  PDFLibrary    : TQuickPDF;
  UnlockResult  : Integer;
  DstPath       : String;
  Rslt          : Integer;
  ErrorCode     : Integer;
begin
  Result := False;
  if Trim(SrcFile) = '' Then
    exit;
  if Trim(DstFile) = '' Then
    exit;

  if not FileExists(Trim(SrcFile)) Then
    exit;
  if not IsFilePDF(SrcFile) then
    exit;
  if not IsFilePDF(DstFile) then
    exit;

  DstPath := j_PathFileName(DstFile);
  if not SysUtils.DirectoryExists(DstPath) then
    SysUtils.ForceDirectories(DstPath);
  if not SysUtils.DirectoryExists(DstPath) then
    exit;

  PDFLibrary := TQuickPDF.Create;
  Try
    UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
    if UnlockResult <> 1 then
      Exit;

    PDFLibrary.LoadFromFile(SrcFile,'');
    Rslt := PDFLibrary.ExtractPages(StartPage, PagesCount);
    if Rslt = 1 Then
    begin
      Rslt := PDFLibrary.SaveToFile(DstFile);
      Result := (Rslt = QuickPdf_NoError);
    end
    else
    begin
      ErrorCode := PDFLibrary.LastErrorCode;
      Try
        AddMsgToEventLog('', J_FirstFileName(ParamStr(0)), 'ExtractFilePages Fail for file [' + SrcFile + '] , ErrorCode = ' + IntToStr(ErrorCode) ,EVENTLOG_ERROR_TYPE, 4, 1);
      Except;
      End;
    end;
  Finally
    FreeAndNil(PDFLibrary);
  End;
end;

end.

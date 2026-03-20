unit dmImage;

interface

{$WARNINGS ON}
{$HINTS ON}
{$WARN UNIT_PLATFORM OFF}
{$WARN SYMBOL_PLATFORM OFF}

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, Db,
  ActiveX, ComObj, ComCtrls, Math, Registry, Jpeg, ShellApi,
  ExtCtrls, olectrls, SHDocVw, mshtml,
  System.IOUtils, IdCoderMIME, System.NetEncoding,
  Word2000, Office2000, MsPpt2000, TypeAndConst,
  FileCtrl, QRExport, Variants, MyUtilities, ApiUtil,
  QuickPdf, EnDiGrph, EnImgScr, EnMisc, EnProLib, EnEncrypt, EnTgaGr,
  EnPngGr, EnPcxGr, EnJpgGr, EnDcxGr, EnTifGr, EnBmpGr, EnIcoGr, EnWmfGr,
  EnTransf, EnDespeckle, EnPrint, NBLib, uAddToPdf, uGlobaltypes,
  htCanvas, htOffice, htRtf, OleServer, Printers, siComp, siLngLnk;

Type
   TFlipType = ( ftHorizontal,ftVertical);

type
  TMinTifFileInfo = Record
    FileName          : String;
    FileExtn          : String;
    Filetype          : String;
    FWidth            : Smallint;
    FHeight           : Smallint;
    FPages            : Smallint;
    FCompression      : TTiffCompression;
    FImageFormat      : TImageFormat;
    TiffOk            : Boolean;
  end;

type
  TWordDoc_Dchiya = Record
    SendDate       : TDateTime;
    OppCompanyName : String;
    NezekDate      : TDateTime;
    OppCarNo       : String;
    CarNo          : String;
    Text           : String;
  end;

type

 Tdm_Image = class(TDataModule)
   PrintDialog: TPrintDialog;
   procedure DataModuleCreate(Sender: TObject);
   procedure DataModuleDestroy(Sender: TObject);
 protected
   { protected declarations }
 private
   { private declarations }
 public
   { Public declarations }

   function  IsWordAppActive : Boolean;
   function  IsExcelAppActive : Boolean;
   function  IsOutlookAppActive : Boolean;
   function  IsAccessAppActive : Boolean;
   function  IsPowerpointAppActive : Boolean;
   function  GetIconFromFile(FileName: String;
                              VCLIcon: TIcon): Boolean;
   function  GetIconAsFile(FileName: String): String;
   function  GetIconAsJpgFile(FileName: String): String;
   function  GetIconAsBMPFile(FileName: String): String;
   procedure CreateIconFromDib(TmpDib: TDibGraphic;
                                IconFileName: String;
                                dpiX, dpiY: Integer);
   function  GetPdfMergeTmpPath : String;
   function  GetTiffInfo(TifFileName : String) : TMinTifFileInfo;
   function  GetTiffColorType(SrcTif: TTiffGraphic) : TTiffColorType;
   function  GetBlackPrecent(CheckFileName : WideString) : Double; OverLoad;
   function  GetBlackPrecent(TmpDib: TDibGraphic) : Double; OverLoad;
   function  GetImageProp(ImageFileName : String) : TImageProp;

   procedure CreateIconBySize(TmpDib: TDibGraphic;
                               IconFileName: String;
                               IconWidth : Integer;
                               IconHeight : Integer);
   procedure ResizeBitmap(Bitmap: TBitmap; const NewWidth, NewHeight: integer);
   procedure TransferGraphicByRation(SrcTif: TTiffGraphic;
                                     DstTif: TTiffGraphic;
                                     fMaxJpgSize, MySize : Integer); OverLoad;
   procedure TransferGraphicByRation(SrcFileName : WideString;
                                     DstTif: TTiffGraphic;
                                     fMaxJpgSize, MySize : Integer); OverLoad;
   function  Create4ImgInTif(TifFileName    : String;
                             ListFileName   : TStringList;
                             IntCompression : Integer;
                             Ratio          : Double) : Boolean;
   procedure DrawStretched(Var TmpDib: TDibGraphic;
                               fWidth: Integer;
                               fHeight: Integer);
   function  SaveTifToFile(TmpTif: TTiffGraphic;
                            ToFileName: TFileName): Boolean;
   function  AppendTifToFile(TmpTif: TTiffGraphic;
                              DstFileName: TFileName;
                              TiffColorType : TTiffColorType): Boolean; OverLoad;
   function  AppendTifToFile(TmpTif: TTiffGraphic;
                              DstFileName: TFileName;
                              DPI_X: Integer;
                              DPI_Y: Integer;
                              TiffColorType : TTiffColorType): Boolean;  OverLoad;
   function  AppendTifToFile(SrcFileName: TFileName;
                             DstFileName: TFileName;
                             TiffColorType : TTiffColorType): Boolean; OverLoad;
   function  AppendTifToFile(SrcFileName: TFileName;
                             DstFileName: TFileName;
                             DPI_X: Integer;
                             DPI_Y: Integer;
                             TiffColorType : TTiffColorType): Boolean; OverLoad;
   function  AddDibToTif(TmpDib: TDibGraphic;
                         TmpTif: TTiffGraphic;
                         TiffColorType : TTiffColorType): Integer;
   procedure FlipImage(SrcDib : TDibGraphic;
                        DstDib : TDibGraphic;
                        FlipType : TFlipType);

   procedure DoGraphicDeskew(SrcDib : TDibGraphic;
                             DstDib : TDibGraphic);
   procedure DoGraphicBorderClean(SrcDib : TDibGraphic;
                                  DstDib : TDibGraphic);
   Procedure DoRotatDIB(TmpDib: TDibGraphic;
                         Angle: Double);
   procedure RotateIfNeed(FileName: String;
                           var Pic: TDibGraphic;
                           Portret: Boolean);
   Procedure DoInvertDIB(TmpDib: TDibGraphic);
   function  IsFileBMP(FileName: String): Boolean;
   function  IsFileJpeg(FileName: String): Boolean;
   function  IsFileTiff(FileName: String): Boolean;
   function  IsFilePNG(FileName: String): Boolean;
   function  IsFilePCX(FileName: String): Boolean;
   function  IsFileWMF(FileName: String): Boolean;
   function  IsFileDCX(FileName: String): Boolean;
   function  IsFilePDF(FileName: String): Boolean;
   function  IsFileDWG(FileName: String): Boolean;
   function  IsFileHTML(FileName: String): Boolean;
   function  IsFileDOC(FileName: String): Boolean;
   function  IsFileDOCX(FileName: String): Boolean;
   function  IsFileWord(FileName : WideString) : Boolean;
   function  IsFileXLS(FileName: String): Boolean;
   function  IsFileXLSX(FileName: String): Boolean;
   function  IsFileExcel(FileName : WideString) : Boolean;
   function  IsFilePPT(FileName: String): Boolean;
   function  IsFilePPTX(FileName: String): Boolean;
   function  IsFilePowerPoint(FileName : WideString) : Boolean;
   function  IsFileOutLookMsg(FileName: String): Boolean;
   function  IsImageShowHTMLViewer(FileName: String): Boolean;
   function  IsImageShowPDFViewer(FileName: String): Boolean;
   function  IsImageShowDWGViewer(FileName: String): Boolean;
   function  IsImageCanAddedToPDF(FileName: String): Boolean;
   function  IsImageShowenInternal(FileName: String): Boolean;
   function  IsImageCanLoadInScan(FileName: String): Boolean;
   function  IsImageDragAndDropSupported(FileName: String): Boolean;
   function  IsItOfficeVersion(FileName: String): Boolean;
   function  IsFileCanBePrinted(FileName: String): Boolean;
   function  IsFileCanBeGraphicPrinted(FileName: String): Boolean;
   function  ImageCanViewByPreView(FileName: String): Boolean;
   function  ImageCanViewByHtmlViewer(FileName: String): Boolean;
   function  CanAddAnnonation(FileName: String): Boolean;
   function  GetImagePageCount(FileName: String): Integer;
   function  GetTifFrameCount(FileName: String): Integer;
   function  CanConvertTo_PDF(FileName: String): Boolean;
   function  CanConvertTo_TIF(FileName: String): Boolean;
   function  CanConvertTo_JPG(FileName: String): Boolean;
   function  CanUpdateFileSize(FileName: String): Boolean;
   function  CanScanMoreToDocument(FileName: String): Boolean;

   function  LoadImageToTif(FileName: String;
                             Var Tmp: TTiffGraphic;
                             SilentMode: Boolean = False;
                             LoadJustFirstPage: Boolean = False;
                             ConvertToBW : Boolean = True): Boolean;
   function  LoadImageToDib(FileName: String;
                            Var Tmp: TDibGraphic;
                            SilentMode: Boolean = False): Boolean;
   function  LoadImageToEnvision(Image: TImage;
                                  Var Tmp: TJpegGraphic): Boolean;
   function  LoadPdfToDib(FileName: String;
                            Var Tmp: TDibGraphic;
                            SilentMode: Boolean = False): Boolean;
   function  BreakTifToPages(FileName: String;
                              ToDir: String;
                              FileLists: TStringList): Boolean;
   function  BreakPdfToTiffPages(PdfFilename : String;
                                  ToDir       : String;
                                  Dpi         : Integer;
                                  FileLists   : TStringList) : Boolean;

   procedure ShowEnvisionInternal(ImageRec: TImageRec);
   procedure ShowGeneralInternal(ImageRec: TImageRec);
   procedure ShowOfficeInternal(ImageRec: TImageRec);
   procedure ShowOutlookInternal(ImageRec: TImageRec);
   procedure ShowPDFInternal(ImageRec: TImageRec);
   procedure ShowHtmlInternal(ImageRec: TImageRec;
                                MailSubject: String);
   procedure ShowImage_Internal(ImageRec: TImageRec); OverLoad;
   procedure ShowImage_Internal(ImageRec: TImageRec;
                                 DoPreview: Boolean); OverLoad;
   procedure ShowImage_External(ImageRec: TImageRec); OverLoad;
   procedure ShowImage_External(FileName: String); OverLoad;

   procedure RemoveDot(Src: TDibGraphic; Dst: TDibGraphic; UseAgresiveMode : Boolean = False; CleanLoop : Integer = 1);
   function  CleanDibBorder(SrcJpgDib: TJpegGraphic): Boolean;

   procedure ConvertWebBrowser2Jpg(const wb: TWebBrowser; const fileName: TFileName) ;
   procedure ConvertDibToGray(SrcDib: TDibGraphic; DstDib: TDibGraphic);
   procedure ConvertDibToGrayLzw(Dst: TDibGraphic);
   procedure ConvertTifToTifBW(SrcTifFile : WideString; DstTifFile : WideString);
   function  ConvertWordToPDF(FileName : String; ToFileName : String) : Boolean;
   function  ConvertWordToRTF(FileName : String; ToFileName : String) : Boolean;
   procedure ConvertPdfToTifBW(SrcPdfFile : WideString; DstTifFile : WideString);
   procedure ConvertDibToBlackWhite(Src: TDibGraphic; Dst: TDibGraphic); OverLoad;
   procedure ConvertDibToBlackWhite(Dst: TDibGraphic); OverLoad;
   procedure ConvertDibGray2BW(Src: TDibGraphic; Dst: TDibGraphic);
   procedure ConvertDib2BW(Src: TDibGraphic; Dst: TDibGraphic);
   procedure ConvertDib2BitMap(Src: TDibGraphic; Dst: TBitMap);
   function  ImgToBase64(FileName : String) : String;
   function  Base64ToImg(Base64Str : String; DibGraphic : TJpegGraphic) : Boolean;
   function  FindMostCleanLine(SrcDib: TDibGraphic;
                               FindCleanLineMode: TFindCleanLineMode;
                               FindCleanLineDirection: TFindCleanLineDirection;
                               CleanLineSearchMethode: TCleanLineSearchMethode;
                               FromPos: Integer): Integer;
   function DoCropDib(SrcDib: TDibGraphic;
                       DstDib: TDibGraphic;
                       CropRect: TRect): Boolean;

   function RenderFileToJpgList(Pdf_FileName: String;
                                ToPath : String;
                                FileList: TStringList): Boolean;
   function RenderPDFToMultiTiff(Pdf_FileName: String;
                                 Tmp: TTiffGraphic): Boolean;
   function RenderFirstPageToTif(Pdf_FileName: String;
                                  Tif_FileName: String): Boolean;
   function RenderPdfPageToStream(FileName : WideString; PageNo : Integer; DPI : Integer; MS : TMemoryStream) : Boolean;
   function RenderFirstPageToDib(Pdf_FileName: String;
                                  Tmp: TDibGraphic): Boolean;
   function ConvertPdfToTif(Pdf_FileName: String;
                             DstFileName: String;
                             ConvertTo_BW: Boolean = True;
                             ConvertJustFirstPage: Boolean = False): Boolean;
   function ConvertPdfToMultiTif(Pdf_FileName: String;
                                  DstFileName: String;
                                  ConvertTo_BW: Boolean = true): Boolean;
   function ConvertTifToPdf(Tif_FileName: String;
                             DstFileName: String): Boolean;
   function ConvertOfficeTo_(Office_FileName: String;
                              Wanted_FileName: String;
                              ToType: Integer;
                              Var ErrMsg: String): Boolean;
   function ConvertPPT_ToRTF(PPT_FileName: String;
                              Wanted_FileName: String;
                              Var ErrMsg: String): Boolean;
   function ConvertRTFToHTML(RTF_FileName: String;
                              Html_FileName: String;
                              Var ErrMsg: String): Boolean;
   function  ConvertOfficeToHtml(OfficeFileName : String; HTMLFileName : String) : Boolean;
   function  ConvertOutlookMsgHtml(MsgFileName : String; HTMLFileName : String) : Boolean;
   function  ConvertOutlookMsg2Html(MsgFileName : String; out HTMLStr : String) : Boolean;
   procedure EncryptBMP(const FileName: String;
                          Key: Integer); OverLoad;
   procedure EncryptBMP(const BMP: TBitmap;
                          Key: Integer); OverLoad;
   procedure EncryptDIB(const FileName: String;
                          Key: Integer); OverLoad;
   procedure EncryptDIB(const TmpDib: TDibGraphic;
                          Key: Integer); OverLoad;

   function  IsOfficeKeyRegistr(Key : WideString) : Boolean;
   function  IsOfficeUserKeyRegistr(Key : WideString) : Boolean;
   procedure FixDibForBarCodeSearch(TmpDib: TBitmapGraphic);

   function  GetYearMonthMethode : String;

   function  GetDocumentMafridType_(DocumentDib: TDibGraphic): TMafridType;  OverLoad;
   function  GetDocumentMafridType_(BarCodesList: TStringList): TMafridType; OverLoad;

   function  IsImageBarCodeScanSupported(FileName: String): Boolean;
   function  GetZXingBarCodesList_(BarCodesList: TStringList;
                                   ClearList: Boolean;
                                   TmpBarCode: TBitmapGraphic): Double; OverLoad;
   function  GetBarCodesList_(BarCodesList: TStringList;
                                ClearList: Boolean;
                                TmpBarCode: TBitmapGraphic): Double; OverLoad;
   function  GetBarCodesList_(BarCodesList: TStringList;
                                ClearList: Boolean;
                                FileName: WideString): Double; OverLoad;

   function  AddImageToPdfFile(SrcFName : String;
                               DstFName : String;
                               FitImage : Boolean): Boolean;
   Function  FindPreviewWindows(OpenFileName: String): Boolean;
   procedure ClosePreviewWindows;
   function  IsPreviewWindowsExists: Boolean;
   procedure PrintManaSticker(BoxNo: String;
                              PackageNo: String;
                              PagesCount: Integer;
                              UserCrt: String; ScanStation: String);
   procedure PrintGraphicDocument(ImageFilename: WideString;
                                   AutoSetOrientation: Boolean); OverLoad;
   procedure PrintGraphicDocument(ImageFilename: WideString;
                                   AutoSetOrientation: Boolean;
                                   DoSelectPrinter: Boolean;
                                   PrintToPrinterName: String); OverLoad;
   procedure JustPrintFile(Filename: WideString;
                            UseDialog : Boolean;
                            DoWait : Boolean = False);
   function  QP_PrintPDF(FileName : WideString) : Integer;
   procedure PrintAnyDocument(ImageFilename: WideString;
                              DoSelectPrinter: Boolean);
   function  GetTextFromDocumentFile(DocFileName : WideString) : WideString;
   function  GetTextFromWordFile(WordFileName : WideString) : WideString;
   function  GetTextFromPDFFile(PdfFileName : WideString) : WideString;
   function  GetTextFromMSGFile(MsgFileName : WideString) : WideString;
   procedure SetTextFromMSGToDb(MzIdent : Integer);
   procedure SetJpgIconToDb(MzDocID : Integer);
   function  ReplaceWordDoc_Dchiya(DataToReplace : TWordDoc_Dchiya;
                                     DocFileName : String;
                                     SaveToFileName : String) : Boolean;
 end;

var
  dm_Image  : Tdm_Image = nil;

Const
  TIF_EXTENTION_NAME: Array [1 .. 2] of String = ('TIF', 'TIFF');
  JPG_EXTENTION_NAME: Array [1 .. 2] of String = ('JPG', 'JPEG');

Const
  Outlook_Proc : WideString = 'outlook.exe';
  Word_Proc    : WideString = 'WinWord.exe';
  Excel_Proc   : WideString = 'Excel.exe';
  PPT_Proc     : WideString = 'Powerpnt.exe';

Const
  FormName_EN      : String = 'Image_InternalPreview_EN';
  FormName_HTML    : String = 'Image_InternalPreview_HTML';
  FormName_PDF     : String = 'Image_InternalPreview_PDF';
  FormName_DWG     : String = 'Image_InternalPreview_DWG';
  FormName_OFFICE  : String = 'Image_InternalPreview_Office';
  FormName_GENERAL : String = 'Image_InternalPreview_General';

(*
To save a file to BLOB:

blob := yourDataset.CreateBlobStream(yourDataset.FieldByName('YOUR_BLOB'), bmWrite);
try
  blob.Seek(0, soFromBeginning);

  fs := TFileStream.Create('c:\your_name.doc', fmOpenRead or fmShareDenyWrite);
  try
    blob.CopyFrom(fs, fs.Size)
  finally
    fs.Free
  end;
finally
  blob.Free
end;

To load from BLOB:

blob := yourDataset.CreateBlobStream(yourDataset.FieldByName('YOUR_BLOB'), bmRead);
try
  blob.Seek(0, soFromBeginning);

  with TFileStream.Create('c:\your_name.doc', fmCreate) do
    try
      CopyFrom(blob, blob.Size)
    finally
      Free
    end;
finally
  blob.Free
end;
*)

implementation
{$R *.DFM}

Uses
  dmJustForSiLang,
  TGeneral,
  rOleTools,
  uFileUtil,
  uPdfUtil,
  OfficeProcess,
  dmWinSoft,
  dmSD_BarCode,
  dmProtect,
  dmDTK_BarCode,
  dmMain,
  dmRo,
  dmOcr,
  UScanPOptForm,
  UScanFrmtForm,
  USelectuserPrinter,
  uViewEnvision,
  uViewOffice,
  uViewOutlook,
  uViewGeneral,
  uViewDwg,
  uViewPDF,
  uViewHTML,
  ZXing.ReadResult,
  ZXing.BarCodeFormat,
  ZXing.ScanManager;

procedure Tdm_Image.DataModuleCreate(Sender: TObject);
begin
  JustWriteTimeToLog('Start - dmImage DataModuleCreate');
  if Application.Terminated then
    Exit;
end;

procedure Tdm_Image.DataModuleDestroy(Sender: TObject);
begin
  Try
    ClosePreviewWindows;
  Except;
  End;
end;

function IsObjectActive(ClassName: string): Boolean;
var
  ClassID: TCLSID;
  Unknown: IUnknown;
begin
  try
    ClassID := ProgIDToClassID(ClassName);
    Result  := GetActiveObject(ClassID, nil, Unknown) = S_OK;
  except
    Result := False;
  end;
end;

function Tdm_Image.IsWordAppActive : Boolean;
begin
  Result := IsObjectActive('Word.Application');
end;

function Tdm_Image.IsExcelAppActive : Boolean;
begin
  Result := IsObjectActive('Excel.Application');
end;

function Tdm_Image.IsOutlookAppActive : Boolean;
begin
  Result := IsObjectActive('Outlook.Application');
end;

function Tdm_Image.IsAccessAppActive : Boolean;
begin
  Result := IsObjectActive('Access.Application');
end;

function Tdm_Image.IsPowerpointAppActive : Boolean;
begin
  Result := IsObjectActive('Powerpoint.Application');
end;

function Tdm_Image.GetIconFromFile(FileName: String; VCLIcon: TIcon): Boolean;
Var
  IconIndex: word;
  Buffer: array[0..2048] of Widechar;
  IconHandle: HIcon;
begin
  Result := False;
  IF Not FileExists(FileName) Then
    Exit;

  StrCopy(@Buffer, PWideChar(FileName));
  IconIndex := 0;
  IconHandle := ExtractAssociatedIcon(HInstance, Buffer, IconIndex);
  if IconHandle <> 0 then
  begin
    VCLIcon.Handle := IconHandle;
    Result := true;
  end;
end;

function Tdm_Image.GetIconAsFile(FileName: String): String;
Var
  VCLIcon: TIcon;
  CurTst: Integer;
  TxtFileName : WideString;
  TxtStr : WideString;
  TmpFile: String;
  Successed: Boolean;
begin
  Result := '';
  VCLIcon := TIcon.Create;
  Try
    IF GetIconFromFile(FileName, VCLIcon) Then
    begin
      Successed := False;
      TmpFile := '';
      For CurTst := 0 To 1000000 Do
      begin
        TmpFile := RemoveBackSlashChar(ParamArea.WinTempDir) + '\Tmp' + IntToStr(CurTst) + '.Ico';
        IF Not FileExists(TmpFile) Then
        begin
          Successed := true;
          Break;
        end;
      end;

      If Successed Then
      begin
        VCLIcon.SaveToFile(TmpFile);
        Result := TmpFile;
      end;
    end
    else
    begin
      TxtFileName := RemoveBackSlashChar(ParamArea.WinTempDir) + '\TmpTxt.Txt';
      TxtStr      := '1';
      WriteTextFile(TxtFileName, TxtStr, True, False);
      IF GetIconFromFile(TxtFileName, VCLIcon) Then
      begin
        Successed := False;
        TmpFile := '';
        For CurTst := 0 To 1000000 Do
        begin
          TmpFile := RemoveBackSlashChar(ParamArea.WinTempDir) + '\Tmp' + IntToStr
            (CurTst) + '.Ico';
          IF Not FileExists(TmpFile) Then
          begin
           Successed := true;
           Break;
          end;
        end;
        If Successed Then
        begin
          VCLIcon.SaveToFile(TmpFile);
          Result := TmpFile;
        end;
        DeleteFile(TxtFileName);
      end
      else
      begin
        TmpFile := RemoveBackSlashChar(ParamArea.WinTempDir) + '\Tmp_9999.Ico';
        NotePadIcon.SaveToFile(TmpFile);
      end;
    end;
  Finally
    FreeAndNil(VCLIcon);
  End;
end;

function Tdm_Image.GetIconAsJpgFile(FileName: String): String;
Var
  IconFileName: String;
  JpgFileName: String;
  JpgDib: TJpegGraphic;
begin
  IconFileName := GetIconAsFile(FileName);
  JpgFileName := ChangeFileExt(IconFileName, '.jpg');
  Result := JpgFileName;

  JpgDib := TJpegGraphic.Create;
  Try
    LoadImageToDib(IconFileName, TDibGraphic(JpgDib), true);
    JpgDib.SaveToFile(Result);
  Finally
    FreeAndNil(JpgDib);
  End;
end;

function Tdm_Image.GetIconAsBMPFile(FileName: String): String;
Var
  IconFileName: String;
  BmpFileName: String;
  BMPDib: TBitmapGraphic;
begin
  IconFileName := GetIconAsFile(FileName);
  BmpFileName := ChangeFileExt(IconFileName, '.BMP');
  Result := BmpFileName;

  BMPDib := TBitmapGraphic.Create;
  Try
    LoadImageToDib(IconFileName, TDibGraphic(BMPDib), true);
    BMPDib.SaveToFile(Result);
  Finally
    FreeAndNil(BMPDib);
  End;
end;

procedure Tdm_Image.CreateIconFromDib(TmpDib: TDibGraphic;
                                        IconFileName: String;
                                        dpiX, dpiY: Integer);
var
  FResizeTransform: TResizeTransform;
  TmpBmp : TBitmapGraphic;
begin
  if (TmpDib.Width = 0) OR (TmpDib.Height = 0) then
    Exit;
  if dpiX < 25 Then
    dpiX := 25;
  if dpiY < 25 Then
    dpiY := 25;

  FResizeTransform := TResizeTransform.Create;
  Try
    TmpBmp := TBitmapGraphic.Create;
    Try
      TmpBmp.Assign(TmpDib);

      if (TmpDib.Width > 64) And (TmpDib.Height > 64) then
      begin
        FResizeTransform.Width := Trunc(TmpBmp.Width * 25 / dpiX);
        FResizeTransform.Height := Trunc(TmpBmp.Height * 25 / dpiY);

        FResizeTransform.Apply(TmpBmp);
      end;

      TmpBmp.SaveToFile(IconFileName);
    Finally
      FreeAndNil(TmpBmp);
    End;
  Finally
   FreeAndNil(FResizeTransform);
  End;
end;

procedure Tdm_Image.CreateIconBySize(TmpDib: TDibGraphic;
                                      IconFileName: String;
                                      IconWidth : Integer;
                                      IconHeight : Integer);
var
  FResizeTransform: TResizeTransform;
  TmpBmp : TBitmapGraphic;
begin
  if (TmpDib.Width = 0) OR (TmpDib.Height = 0) then
    Exit;
  if ((TmpDib.Width < IconWidth) and (TmpDib.Height < IconHeight)) then
    Exit;

  FResizeTransform := TResizeTransform.Create;
  Try
    TmpBmp := TBitmapGraphic.Create;
    Try
      TmpBmp.Assign(TmpDib);
      if ((TmpDib.Width > 64) And (TmpDib.Height > 64)) then
      begin
        FResizeTransform.Width  := IconWidth;
        FResizeTransform.Height := IconHeight;

        FResizeTransform.Apply(TmpBmp);
      end;
      TmpBmp.SaveToFile(IconFileName);
    Finally
      FreeAndNil(TmpBmp);
    End;
  Finally
   FreeAndNil(FResizeTransform);
  End;
end;

function Tdm_Image.GetPdfMergeTmpPath : String;
begin
  Result := RemoveBackSlashChar(ParamArea.WinTempDir) + '\PdfTmp\' + IntToStr(GetTickCount);
  ForceDirectories(Result);
end;

function  Tdm_Image.GetTiffInfo(TifFileName : String) : TMinTifFileInfo;
Var
  CompressTag      : Word;
  CCITTCompressed  : Boolean;
  Stream           : TFileStream;
  BitsPerSample    : Word;
  ImageFormat      : TImageFormat;
  Palette          : TMaxLogPalette;
  MyFileInfo       : TMyFileInfo;
  TmpTif           : TTiffGraphic;
  PhotometricInterpretation  : Word;
begin

  With Result Do
  begin
    FileName      := TifFileName;
    FileExtn      := '';
    Filetype      := '';
    FWidth        := -1;
    FHeight       := -1;
    FPages        := -1;
    FCompression  := tcNotSupported;
    ImageFormat   := ifBlackWhite;
    TiffOk        := False;
  end;

  IF Not FileExists(TifFileName) Then
    Exit;

  MyFileInfo := MyGetFileInfo(TifFileName);
  Try
    Stream := TFileStream.Create(TifFileName,fmOpenRead);
    Try
      TmpTif     := TTiffGraphic.Create;
      Try
        TmpTif.MultiLoad := False;
        TmpTif.LoadFromStream(Stream);
        Result.FileName      := TifFileName;
        Result.FWidth        := TmpTif.Width;
        Result.FHeight       := TmpTif.Height;
        Result.Filetype      := MyFileInfo.TypeName;
        Result.FCompression  := TmpTif.Compression;
        Result.FImageFormat  := TmpTif.ImageFormat;
        Result.FileExtn      := J_ExtFileName(TifFileName);
        Result.TiffOk        := True;
      Finally
        FreeAndNil(TmpTif);
      End;
    Finally
      FreeAndNil(Stream);
    End;
    if Result.TiffOk then
      Result.FPages        := GetImagePageCount(TifFileName);
  Except;
    Result.TiffOk        := False;
  End;
end;

function  Tdm_Image.GetTiffColorType(SrcTif: TTiffGraphic) : TTiffColorType;
begin
  Result := ctBW;
  If Not Assigned(SrcTif) Then
    Exit;
  If SrcTif.IsEmpty Then
    Exit;
  if SrcTif.FrameCount > 0 then
  begin
     IF SrcTif.Frames[1].ImageFormat = ifBlackWhite Then
     begin
       Result := ctBW;
     end
     else
     IF (SrcTif.Frames[1].ImageFormat = ifGray16) or (SrcTif.Frames[1].ImageFormat = ifGray256) Then
     begin
       Result := ctGray;
     end
     else
     begin
       Result := ctColor;
     end;
  end
  else
  begin
    IF SrcTif.ImageFormat = ifBlackWhite Then
    begin
      Result := ctBW;
    end
    else
    IF (SrcTif.ImageFormat = ifGray16) or (SrcTif.ImageFormat = ifGray256) Then
    begin
      Result := ctGray;
    end
    else
    begin
      Result := ctColor;
    end;
  end;
end;

function  Tdm_Image.GetBlackPrecent(CheckFileName : WideString) : Double;
var
  CheckDib : TDibGraphic;
  X : Integer;
  Y : Integer;
  A : TRGB;
  TotalPixcel : Integer;
  BlackPixcel : Integer;
begin
  Result := -1;
  CheckDib := TDibGraphic.Create;
  Try
    TotalPixcel := 0;
    BlackPixcel := 0;
    LoadImageToDib(CheckFileName,CheckDib,True);
    if Not CheckDib.IsEmpty then
    begin
      for X := 0 to CheckDib.Width - 1 do
      begin
        for Y := 0 to CheckDib.Height - 1 do
        begin
          TotalPixcel := TotalPixcel + 1;
          A := CheckDib.Rgb[X, Y];
          if Not (((A.Red > 220) and (A.Blue > 220) and (A.Green > 220)) Or
                  ((A.Red + A.Blue + A.Green) > 600)) Then
          begin
            BlackPixcel := BlackPixcel + 1;
          end { if }
        end; { for }
      end;

      Result := 100 * (BlackPixcel/TotalPixcel);
    end;
  Finally
    FreeAndNil(CheckDib);
  End;
end;

function  Tdm_Image.GetBlackPrecent(TmpDib: TDibGraphic) : Double;
var
  X : Integer;
  Y : Integer;
  A : TRGB;
  TotalPixcel : Integer;
  BlackPixcel : Integer;
begin
  Result := -1;
  Try
    TotalPixcel := 0;
    BlackPixcel := 0;
    if Not TmpDib.IsEmpty then
    begin
      for X := 0 to TmpDib.Width - 1 do
      begin
        for Y := 0 to TmpDib.Height - 1 do
        begin
          TotalPixcel := TotalPixcel + 1;
          A := TmpDib.Rgb[X, Y];
          if Not (((A.Red > 220) and (A.Blue > 220) and (A.Green > 220)) Or
                  ((A.Red + A.Blue + A.Green) > 600)) Then
          begin
            BlackPixcel := BlackPixcel + 1;
          end { if }
        end; { for }
      end;

      Result := 100 * (BlackPixcel/TotalPixcel);
    end;
  Finally
  End;
end;

function  Tdm_Image.GetImageProp(ImageFileName : String) : TImageProp;
Var
  TmpDib : TDibGraphic;
begin
  With Result Do
  begin
    ImgHeight := -1;
    ImgWidth  := -1;
  end;

  if not FileExists(ImageFileName) then
    Exit;

  TmpDib := TDibGraphic.Create;
  Try
    dm_Image.LoadImageToDib(ImageFileName, TmpDib);
    if Not TmpDib.IsEmpty then
    begin
      With Result Do
      begin
        ImgHeight := TmpDib.Height;
        ImgWidth  := TmpDib.Width;
      end;
    end;
  Finally
    FreeAndNil(TmpDib);
  End;
end;

procedure Tdm_Image.ResizeBitmap(Bitmap: TBitmap; const NewWidth, NewHeight: integer);
var
  buffer: TBitmap;
begin
  buffer := TBitmap.Create;
  try
    buffer.SetSize(NewWidth, NewHeight);
    buffer.Canvas.StretchDraw(Rect(0, 0, NewWidth, NewHeight), Bitmap);
    Bitmap.SetSize(NewWidth, NewHeight);
    Bitmap.Canvas.Draw(0, 0, buffer);
  finally
    FreeAndNil(buffer);
  end;
end;

procedure Tdm_Image.TransferGraphicByRation(SrcTif: TTiffGraphic;
                                            DstTif: TTiffGraphic;
                                            fMaxJpgSize, MySize : Integer);
Var
  TmpJpg    : TJpegGraphic;
  SrcDib    : TDibGraphic;
  DstDib    : TDibGraphic;
  Transform : TResizeTransform;
  Ratio     : Double;
  OldWidth  : Integer;
  OldHeight : Integer;
  NewWidth  : Integer;
  NewHeight : Integer;
  CurFrame  : Integer;
  TiffColorType : TTiffColorType;
  NewFileName  : String;
  KeepFileName : String;
begin
  DstTif.ClearAll;
  IF Not SrcTif.IsEmpty Then
  begin
    TiffColorType := GetTiffColorType(SrcTif);
    Ratio  := (fMaxJpgSize/MySize);
    OldWidth  := SrcTif.Width;
    OldHeight := SrcTif.Height;

//
//                               1)       w * h      w2 * h2
//         _w_________________         --------- = --------
//        |          |        |            s1         s2
//        |          |        |
//        |        h2|   s2   |         w * h * s2
//        |          |        |        ------------ = w2 * h2
//        |          |        |             s1
//       h|          |_w2_____|
//        |                   |  2)    w   w2
//        |                   |        - = --
//        |                   |        h   h2
//        |                   |
//        |                   |             (w2 * h)
//        |                   |        h2 = --------
//        |                   |                w
//        |  s1               |
//        |                   |  3)     w * h * s2          (w2 * h)
//        |___________________|        ------------ = w2 * ----------
//                                          s1                 w
//
//                                     w * w * h * s2
//                                     -------------- = w2 * w2
//                                        s1 * h
//
//                                            /-----------------
//                         newWidth =  w2 = \/ (w * w * s2) / s1
//
//                        newHeight =  h2 = (w2 * h) / w
//

    NewWidth := Trunc( Sqrt( OldWidth * OldWidth * Ratio ));
    NewHeight := Trunc((NewWidth * OldHeight) / OldWidth);

    //this make transform closer to fMaxJpgSize
    NewWidth  := Trunc( NewWidth * 0.85 );
    NewHeight := Trunc( NewHeight * 0.85 );
    //this make transform closer to fMaxJpgSize

    if NewWidth < 2 then
        NewWidth := 2;

    if NewHeight < 2 then
        NewHeight := 2;

    IF SrcTif.FrameCount > 0 Then
    begin
      For CurFrame := 1 To SrcTif.FrameCount Do
      begin
        SrcDib    := TDibGraphic.Create;
        DstDib    := TDibGraphic.Create;
        Transform := TResizeTransform.Create;
        Try
          Transform.Width      := NewWidth;
          Transform.Height     := NewHeight;
          SrcDib.Assign(SrcTif.Frames[CurFrame]);

          Transform.ApplyOnDest(SrcDib, DstDib);
          AddDibToTif(DstDib,DstTif,TiffColorType);
        Finally
          FreeAndNil(SrcDib);
          FreeAndNil(DstDib);
          FreeAndNil(Transform);
        End;
      end;
    end
    else
    begin
      Transform := TResizeTransform.Create;
      Try
        Transform.Width      := NewWidth;
        Transform.Height     := NewHeight;

        Transform.ApplyOnDest(SrcTif, DstTif);
      Finally
        FreeAndNil(Transform);
      End;
    end;
  end;
end;

procedure Tdm_Image.TransferGraphicByRation(SrcFileName : WideString;
                                            DstTif : TTiffGraphic;
                                            fMaxJpgSize, MySize : Integer);
Var
  TmpTif    : TTiffGraphic;
  Transform : TResizeTransform;
begin
  DstTif.ClearAll;
  TmpTif := TTiffGraphic.Create;
  Try
    IF LoadImageToTif(SrcFileName,TmpTif,True) Then
    begin
      TransferGraphicByRation(TmpTif,
                              DstTif,
                              fMaxJpgSize, MySize);
    end;
  Finally
    FreeAndNil(TmpTif);
  End;
end;

function  Tdm_Image.Create4ImgInTif(TifFileName    : String;
                                    ListFileName   : TStringList;
                                    IntCompression : Integer;
                                    Ratio          : Double) : Boolean;
Var
  CurFile    : Integer;
  TmpJPG     : TJpegGraphic;
  TmpJPG1    : TJpegGraphic;
  TmpJPG2    : TJpegGraphic;
  TmpJPG3    : TJpegGraphic;
  TmpJPG4    : TJpegGraphic;
  SaveTif    : TTiffGraphic;
  SavePath   : String;
  JpgFile    : String;
  SrcRect    : TRect;
  DstRect    : TRect;
  ErrFound   : Boolean;
  MidHeight  : Integer;
  SecHeight  : Integer;
  NewHeight  : Integer;
  NewWidth   : Integer;
  NewHeight1 : Integer;
  NewWidth1  : Integer;
  NewHeight2 : Integer;
  NewWidth2  : Integer;
  NewHeight3 : Integer;
  NewWidth3  : Integer;
  NewHeight4 : Integer;
  NewWidth4  : Integer;
  NewRatio   : Double;
  OldWidth   : Integer;
  OldHeight  : Integer;
  Transform  : TResizeTransform;

  procedure DoResizeJpg(Jpg : TJpegGraphic);
  Var
    RszJPG : TJpegGraphic;
    NoNeed : Boolean;
  begin
//       See - TransferGraphicByRation
//                                            /-----------------
//                        newWidth  =  w2 = \/ (w * w * s2) / s1
//
//                        newHeight =  h2 = (w2 * h) / w
//
    NoNeed := False;
    if (Jpg.Width <= 640) and (Jpg.Height <= 960) Or
       (Jpg.Width <= 960) and (Jpg.Height <= 640)then
    begin
      NoNeed := True;
    end
    else
    begin
      NewWidth  := Trunc( Sqrt( Jpg.Width * Jpg.Width * Ratio ));
      NewHeight := Trunc((NewWidth * Jpg.Height) / Jpg.Width);
    end;

    If NoNeed Then
      Exit;

    Transform := TResizeTransform.Create;
    Try
      Transform.Width  := NewWidth;
      Transform.Height := NewHeight;

      RszJPG := TJpegGraphic.Create;
      Try
        Transform.ApplyOnDest(Jpg, RszJPG);
        Jpg.Assign(RszJPG);
      Finally
        FreeAndNil(RszJPG);
      End;
    Finally
      FreeAndNil(Transform);
    End;
  end;

  function GetTifCompression(IntCompression : Integer) : TTiffCompression;
  begin
    Result := tcNone;
    case IntCompression of
      1 : Result := tcPackbits;
      2 : Result := tcGroup3_1d;
      3 : Result := tcGroup3_2d;
      4 : Result := tcGroup4;
      5 : Result := tcZLib;
      6 : Result := tcJPEG;
      7 : Result := tcLZW;
    end;
  end;

begin
  Result := False;
  If Not Assigned(ListFileName) Then
    Exit;
  If ListFileName.Count > 4 Then
    Exit;

  SavePath := j_PathFileName(TifFileName);
  If Not DirectoryExists(SavePath) Then
    ForceDirectories(SavePath);
  If Not DirectoryExists(SavePath) Then
    Exit;

  ErrFound := False;
  If FileExists(TifFileName) Then
    DeleteFile(TifFileName);

  SaveTif := TTiffGraphic.Create;
  Try
    if ListFileName.Count = 1 Then
    begin
      JpgFile := ListFileName.Strings[0];
      TmpJPG := TJpegGraphic.Create;
      Try
        TmpJPG.LoadFromFile(JpgFile);
        DoResizeJpg(TmpJPG);
        NewHeight := TmpJPG.Height;
        NewWidth  := TmpJPG.Width;

        SaveTif.NewImage(NewWidth, NewHeight, ifTrueColor, nil, 150, 150);
        SaveTif.Compression :=  GetTifCompression(IntCompression);
        SetStretchBltMode(SaveTif.Canvas.Handle, STRETCH_DELETESCANS);
        SaveTif.Canvas.Brush.Color := clWhite;
        SaveTif.Canvas.FillRect(Rect(0, 0, NewWidth, NewHeight));
        SaveTif.Canvas.Draw(0, 0, SaveTif);

        SrcRect.Top    := 0;
        SrcRect.Left   := 0;
        SrcRect.Right  := TmpJPG.Width;
        SrcRect.Bottom := TmpJPG.Height;

        DstRect.Top    := 0;
        DstRect.Left   := 0;
        DstRect.Right  := TmpJPG.Width;
        DstRect.Bottom := TmpJPG.Height;

        SaveTif.Canvas.CopyRect(DstRect,
                                  TmpJPG.Canvas,
                                  SrcRect);

        SaveTif.SaveToFile(TifFileName);
        ErrFound := False;
      Finally
        FreeAndNil(TmpJPG);
      End;
    end
    else
    if ListFileName.Count = 2 Then
    begin
      // get the Width + height
      TmpJPG1 := TJpegGraphic.Create;
      TmpJPG2 := TJpegGraphic.Create;
      Try
        JpgFile := ListFileName.Strings[0];
        TmpJPG1.LoadFromFile(JpgFile);
        DoResizeJpg(TmpJPG1);
        NewHeight1 := TmpJPG1.Height;
        NewWidth1  := TmpJPG1.Width;

        JpgFile := ListFileName.Strings[1];
        TmpJPG2.LoadFromFile(JpgFile);
        DoResizeJpg(TmpJPG2);
        NewHeight2 := TmpJPG2.Height;
        NewWidth2  := TmpJPG2.Width;

        NewHeight := NewHeight1;
        If NewHeight2 > NewHeight1 Then
          NewHeight := NewHeight2;

        NewWidth  := NewWidth1 + 5 + NewWidth2;

        // create The Image
        SaveTif.NewImage(NewWidth, NewHeight, ifTrueColor, nil, 150, 150);
        SaveTif.Compression :=  GetTifCompression(IntCompression);
        SetStretchBltMode(SaveTif.Canvas.Handle, STRETCH_DELETESCANS);
        SaveTif.Canvas.Brush.Color := clWhite;
        SaveTif.Canvas.FillRect(Rect(0, 0, NewWidth, NewHeight));
        SaveTif.Canvas.Draw(0, 0, SaveTif);

        TmpJPG := TJpegGraphic.Create;
        Try
          TmpJPG.Assign(TmpJPG1);

          SrcRect.Top    := 0;
          SrcRect.Left   := 0;
          SrcRect.Right  := TmpJPG.Width;
          SrcRect.Bottom := TmpJPG.Height;

          DstRect.Top    := 0;
          DstRect.Left   := 0;
          DstRect.Right  := TmpJPG.Width;
          DstRect.Bottom := TmpJPG.Height;

          SaveTif.Canvas.CopyRect(DstRect,
                                    TmpJPG.Canvas,
                                    SrcRect);
        Finally
          FreeAndNil(TmpJPG);
        End;

        TmpJPG := TJpegGraphic.Create;
        Try
          TmpJPG.Assign(TmpJPG2);

          SrcRect.Top    := 0;
          SrcRect.Left   := 0;
          SrcRect.Right  := TmpJPG.Width;
          SrcRect.Bottom := TmpJPG.Height;

          DstRect.Top    := 0;
          DstRect.Left   := NewWidth1 + 5;
          DstRect.Right  := SaveTif.Width;
          DstRect.Bottom := TmpJPG.Height;

          SaveTif.Canvas.CopyRect(DstRect,
                                    TmpJPG.Canvas,
                                    SrcRect);
        Finally
          FreeAndNil(TmpJPG);
        End;
      Finally
        FreeAndNil(TmpJPG1);
        FreeAndNil(TmpJPG2);
      End;

      SaveTif.SaveToFile(TifFileName);
      ErrFound := False;
    end
    else
    if ListFileName.Count = 3 Then
    begin
      // get the Width + height
      TmpJPG1 := TJpegGraphic.Create;
      TmpJPG2 := TJpegGraphic.Create;
      TmpJPG3 := TJpegGraphic.Create;
      Try
        JpgFile := ListFileName.Strings[0];
        TmpJPG1.LoadFromFile(JpgFile);
        DoResizeJpg(TmpJPG1);
        NewHeight1 := TmpJPG1.Height;
        NewWidth1  := TmpJPG1.Width;

        JpgFile := ListFileName.Strings[1];
        TmpJPG2.LoadFromFile(JpgFile);
        DoResizeJpg(TmpJPG2);
        NewHeight2 := TmpJPG2.Height;
        NewWidth2  := TmpJPG2.Width;

        JpgFile := ListFileName.Strings[2];
        TmpJPG3.LoadFromFile(JpgFile);
        DoResizeJpg(TmpJPG3);
        NewHeight3 := TmpJPG3.Height;
        NewWidth3  := TmpJPG3.Width;

        NewHeight := NewHeight1;
        If NewHeight2 > NewHeight1 Then
          NewHeight := NewHeight2;

        MidHeight := NewHeight;

        NewHeight := NewHeight + 5 + NewHeight3;

        NewWidth  := NewWidth1 + 5 + NewWidth2;
        if NewWidth3 > NewWidth Then
          NewWidth  := NewWidth3;

        // create The Image
        SaveTif.NewImage(NewWidth, NewHeight, ifTrueColor, nil, 150, 150);
        SaveTif.Compression :=  GetTifCompression(IntCompression);
        SetStretchBltMode(SaveTif.Canvas.Handle, STRETCH_DELETESCANS);
        SaveTif.Canvas.Brush.Color := clWhite;
        SaveTif.Canvas.FillRect(Rect(0, 0, NewWidth, NewHeight));
        SaveTif.Canvas.Draw(0, 0, SaveTif);

        TmpJPG := TJpegGraphic.Create;
        Try
          TmpJPG.Assign(TmpJPG1);

          SrcRect.Top    := 0;
          SrcRect.Left   := 0;
          SrcRect.Right  := TmpJPG.Width;
          SrcRect.Bottom := TmpJPG.Height;

          DstRect.Top    := 0;
          DstRect.Left   := 0;
          DstRect.Right  := TmpJPG.Width;
          DstRect.Bottom := TmpJPG.Height;

          SaveTif.Canvas.CopyRect(DstRect,
                                    TmpJPG.Canvas,
                                    SrcRect);
        Finally
          FreeAndNil(TmpJPG);
        End;

        TmpJPG := TJpegGraphic.Create;
        Try
          TmpJPG.Assign(TmpJPG2);

          SrcRect.Top    := 0;
          SrcRect.Left   := 0;
          SrcRect.Right  := TmpJPG.Width;
          SrcRect.Bottom := TmpJPG.Height;

          DstRect.Top    := 0;
          DstRect.Left   := NewWidth1 + 5;
          DstRect.Right  := SaveTif.Width;
          DstRect.Bottom := TmpJPG.Height;

          SaveTif.Canvas.CopyRect(DstRect,
                                    TmpJPG.Canvas,
                                    SrcRect);
        Finally
          FreeAndNil(TmpJPG);
        End;

        TmpJPG := TJpegGraphic.Create;
        Try
          TmpJPG.Assign(TmpJPG3);

          SrcRect.Top    := 0;
          SrcRect.Left   := 0;
          SrcRect.Right  := TmpJPG.Width;
          SrcRect.Bottom := TmpJPG.Height;

          DstRect.Top    := MidHeight + 5;
          DstRect.Left   := 0;
          DstRect.Right  := TmpJPG.Width;
          DstRect.Bottom := SaveTif.Height;

          SaveTif.Canvas.CopyRect(DstRect,
                                    TmpJPG.Canvas,
                                    SrcRect);
        Finally
          FreeAndNil(TmpJPG);
        End;
      Finally
        FreeAndNil(TmpJPG1);
        FreeAndNil(TmpJPG2);
        FreeAndNil(TmpJPG3);
      End;

      SaveTif.SaveToFile(TifFileName);
      ErrFound := False;
    end
    else
    if ListFileName.Count = 4 Then
    begin
      // get the Width + height
      TmpJPG1 := TJpegGraphic.Create;
      TmpJPG2 := TJpegGraphic.Create;
      TmpJPG3 := TJpegGraphic.Create;
      TmpJPG4 := TJpegGraphic.Create;
      Try
        JpgFile := ListFileName.Strings[0];
        TmpJPG1.LoadFromFile(JpgFile);
        DoResizeJpg(TmpJPG1);
        NewHeight1 := TmpJPG1.Height;
        NewWidth1  := TmpJPG1.Width;

        JpgFile := ListFileName.Strings[1];
        TmpJPG2.LoadFromFile(JpgFile);
        DoResizeJpg(TmpJPG2);
        NewHeight2 := TmpJPG2.Height;
        NewWidth2  := TmpJPG2.Width;

        JpgFile := ListFileName.Strings[2];
        TmpJPG3.LoadFromFile(JpgFile);
        DoResizeJpg(TmpJPG3);
        NewHeight3 := TmpJPG3.Height;
        NewWidth3  := TmpJPG3.Width;

        JpgFile := ListFileName.Strings[3];
        TmpJPG4.LoadFromFile(JpgFile);
        DoResizeJpg(TmpJPG4);
        NewHeight4 := TmpJPG4.Height;
        NewWidth4  := TmpJPG4.Width;

        NewHeight := NewHeight1;
        If NewHeight2 > NewHeight1 Then
          NewHeight := NewHeight2;

        MidHeight := NewHeight;

        SecHeight := NewHeight3;
        If NewHeight4 > NewHeight3 Then
          SecHeight := NewHeight4;

        NewHeight := NewHeight + 5 + SecHeight;

        NewWidth  := NewWidth1 + 5 + NewWidth2;
        if NewWidth3 + 5 + NewWidth4 > NewWidth Then
          NewWidth  := NewWidth3 + 5 + NewWidth4;

        // create The Image
        SaveTif.NewImage(NewWidth, NewHeight, ifTrueColor, nil, 200, 200);
        SaveTif.Compression :=  GetTifCompression(IntCompression);
        SetStretchBltMode(SaveTif.Canvas.Handle, STRETCH_DELETESCANS);
        SaveTif.Canvas.Brush.Color := clWhite;
        SaveTif.Canvas.FillRect(Rect(0, 0, NewWidth, NewHeight));
        SaveTif.Canvas.Draw(0, 0, SaveTif);

        TmpJPG := TJpegGraphic.Create;
        Try
          TmpJPG.Assign(TmpJPG1);

          SrcRect.Top    := 0;
          SrcRect.Left   := 0;
          SrcRect.Right  := TmpJPG.Width;
          SrcRect.Bottom := TmpJPG.Height;

          DstRect.Top    := 0;
          DstRect.Left   := 0;
          DstRect.Right  := TmpJPG.Width;
          DstRect.Bottom := TmpJPG.Height;

          SaveTif.Canvas.CopyRect(DstRect,
                                    TmpJPG.Canvas,
                                    SrcRect);
        Finally
          FreeAndNil(TmpJPG);
        End;

        TmpJPG := TJpegGraphic.Create;
        Try
          TmpJPG.Assign(TmpJPG2);

          SrcRect.Top    := 0;
          SrcRect.Left   := 0;
          SrcRect.Right  := TmpJPG.Width;
          SrcRect.Bottom := TmpJPG.Height;

          DstRect.Top    := 0;
          DstRect.Left   := NewWidth1 + 5;
          DstRect.Right  := NewWidth1 + 5 + TmpJPG.Width;
          DstRect.Bottom := TmpJPG.Height;

          SaveTif.Canvas.CopyRect(DstRect,
                                    TmpJPG.Canvas,
                                    SrcRect);
        Finally
          FreeAndNil(TmpJPG);
        End;

        TmpJPG := TJpegGraphic.Create;
        Try
          TmpJPG.Assign(TmpJPG3);

          SrcRect.Top    := 0;
          SrcRect.Left   := 0;
          SrcRect.Right  := TmpJPG.Width;
          SrcRect.Bottom := TmpJPG.Height;

          DstRect.Top    := MidHeight + 5;
          DstRect.Left   := 0;
          DstRect.Right  := TmpJPG.Width;
          DstRect.Bottom := MidHeight + 5 + TmpJPG.Height;

          SaveTif.Canvas.CopyRect(DstRect,
                                    TmpJPG.Canvas,
                                    SrcRect);
        Finally
          FreeAndNil(TmpJPG);
        End;

        TmpJPG := TJpegGraphic.Create;
        Try
          TmpJPG.Assign(TmpJPG4);

          SrcRect.Top    := 0;
          SrcRect.Left   := 0;
          SrcRect.Right  := TmpJPG.Width;
          SrcRect.Bottom := TmpJPG.Height;

          DstRect.Top    := MidHeight + 5;
          DstRect.Left   := NewWidth3 + 5;
          DstRect.Right  := NewWidth3 + 5 + TmpJPG.Width;
          DstRect.Bottom := MidHeight + 5 + TmpJPG.Height;

          SaveTif.Canvas.CopyRect(DstRect,
                                    TmpJPG.Canvas,
                                    SrcRect);
        Finally
          FreeAndNil(TmpJPG);
        End;
      Finally
        FreeAndNil(TmpJPG1);
        FreeAndNil(TmpJPG2);
        FreeAndNil(TmpJPG3);
        FreeAndNil(TmpJPG4);
      End;

      SaveTif.SaveToFile(TifFileName);
      ErrFound := False;
    end;
  Finally
    FreeAndNil(SaveTif);
  End;

  Result := Not ErrFound;
end;

procedure Tdm_Image.DrawStretched(Var TmpDib: TDibGraphic;
                                      fWidth: Integer;
                                      fHeight: Integer);
var
  Factor: Single;
  FResizeTransform: TResizeTransform;
begin
  if (TmpDib = nil) Or (TmpDib.IsEmpty) Or (TmpDib.Width = 0) Or
    (TmpDib.Height = 0) Or (fWidth = 0) Or (fHeight = 0) Then
   Exit;

  if (TmpDib.Width < fWidth) and (TmpDib.Height < fHeight) then
   Exit
  else if TmpDib.Width > TmpDib.Height then
   Factor := fWidth / TmpDib.Width
  else
   Factor := fHeight / TmpDib.Height;

  FResizeTransform := TResizeTransform.Create;
  Try
   FResizeTransform.Width := SafeTrunc(TmpDib.Width * Factor);
   FResizeTransform.Height := SafeTrunc(TmpDib.Height * Factor);

   if FResizeTransform.Width < 2 then
    FResizeTransform.Width := 2;

   if FResizeTransform.Height < 2 then
    FResizeTransform.Height := 2;

   FResizeTransform.Apply(TmpDib);
  Finally
   FreeAndNil(FResizeTransform);
  End;
end;

function Tdm_Image.SaveTifToFile(TmpTif: TTiffGraphic;
                                  ToFileName: TFileName): Boolean;
Var
  CurFrame: Integer;
  SrcRect: TRect;
  DstRect: TRect;
  Stream: TFileStream;
  SaveTif: TTiffGraphic;
  DumyTif: TTiffGraphic;
  TifImgFormat : TImageFormat;
  SavePath: String;
begin
  Result := False;
  IF TmpTif.IsEmpty Then
   Exit;

  TifImgFormat := TmpTif.ImageFormat;

  IF TmpTif.FrameCount = 0 Then
  begin
    IF FileExists(ToFileName) Then
      DeleteFile(ToFileName);

    SavePath := j_PathFileName(ToFileName);
    If Not DirectoryExists(SavePath) Then
      ForceDirectories(SavePath);
    SaveTif := TTiffGraphic.Create;
    Stream := TFileStream.Create(ToFileName, fmCreate { fmOpenReadWrite } );
    Try
      SaveTif.ClearAll;
      IF TifImgFormat = ifBlackWhite Then
      begin
        SaveTif.NewImage(TmpTif.Width, TmpTif.Height, ifBlackWhite, nil,
          TmpTif.XDotsPerInch, TmpTif.YDotsPerInch);
        SaveTif.Compression := tcGroup4;
      end
      else
      IF (TifImgFormat = ifGray16) or (TifImgFormat = ifGray256) Then
      begin
        SaveTif.NewImage(TmpTif.Width, TmpTif.Height, ifGray256, nil,
          TmpTif.XDotsPerInch, TmpTif.YDotsPerInch);
        SaveTif.Compression := tcJPEG;
      end
      else
      begin
       SaveTif.NewImage(TmpTif.Width, TmpTif.Height, ifTrueColor, nil,
        TmpTif.XDotsPerInch, TmpTif.YDotsPerInch);
       SaveTif.Compression := { tcJPEG } tcLZW;
      end;
      SetStretchBltMode(SaveTif.Canvas.Handle, STRETCH_DELETESCANS);

      SrcRect.Top := 0;
      SrcRect.Left := 0;
      SrcRect.Right := TmpTif.Width;
      SrcRect.Bottom := TmpTif.Height;

      DstRect.Top := 0;
      DstRect.Left := 0;
      DstRect.Right := TmpTif.Width;
      DstRect.Bottom := TmpTif.Height;

      SaveTif.Canvas.CopyRect(DstRect, TmpTif.Canvas, SrcRect);
      SaveTif.SaveToStream(Stream);
    Finally
      FreeAndNil(Stream);
      FreeAndNil(SaveTif);
    End;
    Result := true;
  end
  else
  begin
    IF FileExists(ToFileName) Then
     DeleteFile(ToFileName);

    SavePath := j_PathFileName(ToFileName);
    If Not DirectoryExists(SavePath) Then
      ForceDirectories(SavePath);
    SaveTif := TTiffGraphic.Create;
    Stream := TFileStream.Create(ToFileName, fmCreate { fmOpenReadWrite } );
    Try
      For CurFrame := 1 To TmpTif.FrameCount Do
      begin
        TifImgFormat := TmpTif.Frames[CurFrame].ImageFormat;
        SaveTif.ClearAll;
        IF TifImgFormat = ifBlackWhite Then
        begin
          SaveTif.NewImage(TmpTif.Frames[CurFrame].Width,
            TmpTif.Frames[CurFrame].Height, ifBlackWhite, nil,
            TmpTif.Frames[CurFrame].XDotsPerInch,
            TmpTif.Frames[CurFrame].YDotsPerInch);
          SaveTif.Compression := tcGroup4;
        end
        else
        IF (TifImgFormat = ifGray16) or (TifImgFormat = ifGray256) Then
        begin
          SaveTif.NewImage(TmpTif.Frames[CurFrame].Width,
            TmpTif.Frames[CurFrame].Height, ifGray256, nil,
            TmpTif.Frames[CurFrame].XDotsPerInch,
            TmpTif.Frames[CurFrame].YDotsPerInch);
          SaveTif.Compression := tcJPEG;
        end
        else
        begin
          SaveTif.NewImage(TmpTif.Frames[CurFrame].Width,
            TmpTif.Frames[CurFrame].Height, ifTrueColor, nil,
            TmpTif.Frames[CurFrame].XDotsPerInch,
            TmpTif.Frames[CurFrame].YDotsPerInch);
          SaveTif.Compression := { tcJPEG } tcLZW;
        end;
        SetStretchBltMode(SaveTif.Canvas.Handle, STRETCH_DELETESCANS);

        SrcRect.Top := 0;
        SrcRect.Left := 0;
        SrcRect.Right := TmpTif.Frames[CurFrame].Width;
        SrcRect.Bottom := TmpTif.Frames[CurFrame].Height;

        DstRect.Top := 0;
        DstRect.Left := 0;
        DstRect.Right := TmpTif.Frames[CurFrame].Width;
        DstRect.Bottom := TmpTif.Frames[CurFrame].Height;

        SaveTif.Canvas.CopyRect(DstRect, TmpTif.Frames[CurFrame].Canvas, SrcRect);
        IF CurFrame = 1 Then
          SaveTif.SaveToStream(Stream)
        else
          SaveTif.AppendToStream(Stream);
      end;
    Finally
      FreeAndNil(Stream);
      FreeAndNil(SaveTif);
    End;
   Result := true;
  end;
end;

function Tdm_Image.AppendTifToFile(TmpTif: TTiffGraphic;
                                    DstFileName: TFileName;
                                    TiffColorType : TTiffColorType): Boolean;
Var
  CurFrame: Integer;
  SrcRect: TRect;
  DstRect: TRect;
  Stream: TFileStream;
  SaveTif: TTiffGraphic;
  SavePath: String;
begin
  Result := False;

  IF TmpTif.IsEmpty Then
    Exit;

  SavePath := j_PathFileName(DstFileName);
  If Not DirectoryExists(SavePath) Then
    ForceDirectories(SavePath);
  IF Not FileExists(DstFileName) Then
    Stream := TFileStream.Create(DstFileName, fmCreate)
  else
    Stream := TFileStream.Create(DstFileName, fmOpenReadWrite);

  Try
    IF TmpTif.FrameCount = 0 Then
    begin
      SaveTif := TTiffGraphic.Create;
      Try
        SaveTif.ClearAll;
        IF TiffColorType = ctBW Then
        begin
          SaveTif.NewImage(TmpTif.Width, TmpTif.Height, ifBlackWhite, nil,
            TmpTif.XDotsPerInch, TmpTif.YDotsPerInch);
          SaveTif.Compression := tcGroup4;
        end
        else
        IF TiffColorType = ctGray Then
        begin
          SaveTif.NewImage(TmpTif.Width, TmpTif.Height, ifGray256, nil,
            TmpTif.XDotsPerInch, TmpTif.YDotsPerInch);
          SaveTif.Compression := tcJPEG;
        end
        else
        begin
          SaveTif.NewImage(TmpTif.Width, TmpTif.Height, ifTrueColor, nil,
            TmpTif.XDotsPerInch, TmpTif.YDotsPerInch);
          SaveTif.Compression := { tcJPEG } tcLZW;
        end;
        SetStretchBltMode(SaveTif.Canvas.Handle, STRETCH_DELETESCANS);

        SrcRect.Top := 0;
        SrcRect.Left := 0;
        SrcRect.Right := TmpTif.Width;
        SrcRect.Bottom := TmpTif.Height;

        DstRect.Top := 0;
        DstRect.Left := 0;
        DstRect.Right := TmpTif.Width;
        DstRect.Bottom := TmpTif.Height;

        SaveTif.Canvas.CopyRect(DstRect, TmpTif.Canvas, SrcRect);

        IF Not FileExists(DstFileName) Then
          SaveTif.SaveToStream(Stream)
        else
          SaveTif.AppendToStream(Stream);
      Finally
        FreeAndNil(SaveTif);
      End;
      Result := true;
    end
    else
    begin
      SaveTif := TTiffGraphic.Create;
      Try
        For CurFrame := 1 To TmpTif.FrameCount Do
        begin
          SaveTif.ClearAll;
          IF TiffColorType = ctBW Then
          begin
            SaveTif.NewImage(TmpTif.Frames[CurFrame].Width,
              TmpTif.Frames[CurFrame].Height, ifBlackWhite, nil,
              TmpTif.Frames[CurFrame].XDotsPerInch,
              TmpTif.Frames[CurFrame].YDotsPerInch);
            SaveTif.Compression := tcGroup4;
          end
          else
          IF TiffColorType = ctGray Then
          begin
            SaveTif.NewImage(TmpTif.Frames[CurFrame].Width,
              TmpTif.Frames[CurFrame].Height, ifGray256, nil,
              TmpTif.Frames[CurFrame].XDotsPerInch,
              TmpTif.Frames[CurFrame].YDotsPerInch);
            SaveTif.Compression := tcJPEG;
          end
          else
          begin
            SaveTif.NewImage(TmpTif.Frames[CurFrame].Width,
              TmpTif.Frames[CurFrame].Height, ifTrueColor, nil,
              TmpTif.Frames[CurFrame].XDotsPerInch,
              TmpTif.Frames[CurFrame].YDotsPerInch);
            SaveTif.Compression := { tcJPEG } tcLZW;
          end;
          SetStretchBltMode(SaveTif.Canvas.Handle, STRETCH_DELETESCANS);

          SrcRect.Top := 0;
          SrcRect.Left := 0;
          SrcRect.Right := TmpTif.Frames[CurFrame].Width;
          SrcRect.Bottom := TmpTif.Frames[CurFrame].Height;

          DstRect.Top := 0;
          DstRect.Left := 0;
          DstRect.Right := TmpTif.Frames[CurFrame].Width;
          DstRect.Bottom := TmpTif.Frames[CurFrame].Height;

          SaveTif.Canvas.CopyRect(DstRect, TmpTif.Frames[CurFrame].Canvas, SrcRect);

          IF Not FileExists(DstFileName) Then
            SaveTif.SaveToStream(Stream)
          else
            SaveTif.AppendToStream(Stream);
        end;
      Finally
       FreeAndNil(SaveTif);
      End;
      Result := true;
    end;
  Finally
   FreeAndNil(Stream);
  End;
end;

function Tdm_Image.AppendTifToFile(TmpTif: TTiffGraphic;
                                     DstFileName: TFileName;
                                     DPI_X: Integer;
                                     DPI_Y: Integer;
                                     TiffColorType : TTiffColorType): Boolean;
Var
  CurFrame: Integer;
  SrcRect: TRect;
  DstRect: TRect;
  Stream: TFileStream;
  SaveTif: TTiffGraphic;
  SavePath: String;
begin
  Result := False;

  IF DPI_X < 100 Then
    DPI_X := 100;
  IF DPI_X > 1200 Then
    DPI_X := 1200;
  IF DPI_Y < 100 Then
    DPI_Y := 100;
  IF DPI_Y > 1200 Then
    DPI_Y := 1200;

  IF TmpTif.IsEmpty Then
    Exit;

  SavePath := j_PathFileName(DstFileName);
  If Not DirectoryExists(SavePath) Then
    ForceDirectories(SavePath);
  IF Not FileExists(DstFileName) Then
    Stream := TFileStream.Create(DstFileName, fmCreate)
  else
    Stream := TFileStream.Create(DstFileName, fmOpenReadWrite);

  Try
    IF TmpTif.FrameCount = 0 Then
    begin
      SaveTif := TTiffGraphic.Create;
      Try
        SaveTif.ClearAll;
        IF TiffColorType = ctBW Then
        begin
          SaveTif.NewImage(TmpTif.Width, TmpTif.Height, ifBlackWhite, nil, DPI_X, DPI_Y);
          SaveTif.Compression := tcGroup4;
        end
        else
        IF TiffColorType = ctGray Then
        begin
          SaveTif.NewImage(TmpTif.Width, TmpTif.Height, ifGray256, nil, DPI_X, DPI_Y);
          SaveTif.Compression := tcJPEG;
        end
        else
        begin
          SaveTif.NewImage(TmpTif.Width, TmpTif.Height, ifTrueColor, nil, DPI_X, DPI_Y);
          SaveTif.Compression := { tcJPEG } tcLZW;
        end;
        SetStretchBltMode(SaveTif.Canvas.Handle, STRETCH_DELETESCANS);

        SrcRect.Top := 0;
        SrcRect.Left := 0;
        SrcRect.Right := TmpTif.Width;
        SrcRect.Bottom := TmpTif.Height;

        DstRect.Top := 0;
        DstRect.Left := 0;
        DstRect.Right := TmpTif.Width;
        DstRect.Bottom := TmpTif.Height;

        SaveTif.Canvas.CopyRect(DstRect, TmpTif.Canvas, SrcRect);

        IF Not FileExists(DstFileName) Then
         SaveTif.SaveToStream(Stream)
        else
         SaveTif.AppendToStream(Stream);
      Finally
        FreeAndNil(SaveTif);
      End;
      Result := true;
     end
     else
     begin
      SaveTif := TTiffGraphic.Create;
      Try
        For CurFrame := 1 To TmpTif.FrameCount Do
        begin
          SaveTif.ClearAll;
          IF TiffColorType = ctBW Then
          begin
            SaveTif.NewImage(TmpTif.Frames[CurFrame].Width,
              TmpTif.Frames[CurFrame].Height, ifBlackWhite, nil, DPI_X, DPI_Y);
            SaveTif.Compression := tcGroup4;
          end
          else
          IF TiffColorType = ctGray Then
          begin
            SaveTif.NewImage(TmpTif.Frames[CurFrame].Width,
              TmpTif.Frames[CurFrame].Height, ifGray256, nil, DPI_X, DPI_Y);
            SaveTif.Compression := tcJPEG;
          end
          else
          begin
            SaveTif.NewImage(TmpTif.Frames[CurFrame].Width,
              TmpTif.Frames[CurFrame].Height, ifTrueColor, nil, DPI_X, DPI_Y);
            SaveTif.Compression := { tcJPEG } tcLZW;
          end;
          SetStretchBltMode(SaveTif.Canvas.Handle, STRETCH_DELETESCANS);

          SrcRect.Top := 0;
          SrcRect.Left := 0;
          SrcRect.Right := TmpTif.Frames[CurFrame].Width;
          SrcRect.Bottom := TmpTif.Frames[CurFrame].Height;

          DstRect.Top := 0;
          DstRect.Left := 0;
          DstRect.Right := TmpTif.Frames[CurFrame].Width;
          DstRect.Bottom := TmpTif.Frames[CurFrame].Height;

          SaveTif.Canvas.CopyRect(DstRect, TmpTif.Frames[CurFrame].Canvas, SrcRect);

          IF Not FileExists(DstFileName) Then
            SaveTif.SaveToStream(Stream)
          else
            SaveTif.AppendToStream(Stream);
        end;
      Finally
       FreeAndNil(SaveTif);
      End;
      Result := true;
    end;
  Finally
    FreeAndNil(Stream);
  End;
end;

function Tdm_Image.AppendTifToFile(SrcFileName: TFileName;
                                     DstFileName: TFileName;
                                     TiffColorType : TTiffColorType): Boolean;
Var
  TmpTif: TTiffGraphic;
begin
  Result := False;
  if not FileExists(SrcFileName) then
    Exit;
  TmpTif := TTiffGraphic.Create;
  TmpTif.MultiLoad := true;
  Try
    if IsFileTiff(SrcFileName) then
      TmpTif.LoadFromFile(SrcFileName)
    else
      LoadImageToTif(SrcFileName,Tmptif,False{Silent},False {Loadjustfirst},False {ConvertToBW});

    IF TmpTif.IsEmpty Then
     Exit;
    Result := AppendTifToFile(TmpTif, DstFileName, TiffColorType);
  Finally
   FreeAndNil(TmpTif);
  End;
end;

function Tdm_Image.AppendTifToFile(SrcFileName: TFileName;
                                     DstFileName: TFileName;
                                     DPI_X: Integer; DPI_Y: Integer;
                                     TiffColorType : TTiffColorType): Boolean;
Var
  TmpTif: TTiffGraphic;
begin
  Result := False;

  TmpTif := TTiffGraphic.Create;
  TmpTif.MultiLoad := true;
  Try
   TmpTif.LoadFromFile(SrcFileName);

   IF TmpTif.IsEmpty Then
    Exit;
   Result := AppendTifToFile(TmpTif, DstFileName, DPI_X, DPI_Y, TiffColorType);
  Finally
   FreeAndNil(TmpTif);
  End;
end;

function Tdm_Image.AddDibToTif(TmpDib: TDibGraphic;
                               TmpTif: TTiffGraphic;
                               TiffColorType : TTiffColorType): Integer;
Var
  SrcRect: TRect;
  DstRect: TRect;
  CurFrame: Integer;
  FileStream: TFileStream;
  DoAppend: Boolean;
  TmpFileName: String;
  SavePath: String;
begin
  Result := -1;
  IF TmpDib.IsEmpty Then
   Exit;

  TmpFileName := RemoveBackSlashChar(ParamArea.WinTempDir) + '\TmpTif.Tif';
  IF FileExists(TmpFileName) Then
    DeleteFile(TmpFileName);
  Try
    SavePath := j_PathFileName(TmpFileName);
    If Not DirectoryExists(SavePath) Then
     ForceDirectories(SavePath);
    FileStream := TFileStream.Create(TmpFileName, fmCreate { fmOpenReadWrite } );
    Try
      DoAppend := False;
      IF Not TmpTif.IsEmpty Then
      begin
        DoAppend := true;
        IF TmpTif.FrameCount = 0 Then
        begin
          TmpTif.SaveToStream(FileStream)
        end
        else
        for CurFrame := 1 to TmpTif.FrameCount do
        begin
         if CurFrame = 1 then
          TmpTif.Frames[CurFrame].SaveToStream(FileStream)
         else
          TmpTif.Frames[CurFrame].AppendToStream(FileStream);
        end;
      end;

      TmpTif.ClearAll;
      If TiffColorType = ctBW Then
      begin
        TmpTif.NewImage(TmpDib.Width, TmpDib.Height, ifBlackWhite, nil,
          TmpDib.XDotsPerInch, TmpDib.YDotsPerInch);
        TmpTif.Compression := tcGroup4;
      end
      else
      If TiffColorType = ctGray Then
      begin
        TmpTif.NewImage(TmpDib.Width, TmpDib.Height, ifGray256, nil,
          TmpDib.XDotsPerInch, TmpDib.YDotsPerInch);
        TmpTif.Compression := tcJPEG;
      end
      else
      begin
        TmpTif.NewImage(TmpDib.Width, TmpDib.Height, ifTrueColor, nil,
          TmpDib.XDotsPerInch, TmpDib.YDotsPerInch);
        TmpTif.Compression := { tcJPEG } tcLZW;
      end;
      SetStretchBltMode(TmpTif.Canvas.Handle, STRETCH_DELETESCANS);

      SrcRect.Top := 0;
      SrcRect.Left := 0;
      SrcRect.Right := TmpDib.Width;
      SrcRect.Bottom := TmpDib.Height;

      DstRect.Top := 0;
      DstRect.Left := 0;
      DstRect.Right := TmpDib.Width;
      DstRect.Bottom := TmpDib.Height;

      TmpTif.Canvas.CopyRect(DstRect, TmpDib.Canvas, SrcRect);

      IF Not DoAppend Then
      begin
        TmpTif.SaveToStream(FileStream)
      end
      else
      begin
        TmpTif.AppendToStream(FileStream);
      end;
    Finally
      FreeAndNil(FileStream);
    End;
    TmpTif.MultiLoad := true;
    TmpTif.LoadFromFile(TmpFileName);
  Finally
   IF FileExists(TmpFileName) Then
    DeleteFile(TmpFileName);
  End;
  Result := TmpTif.FrameCount;
end;

procedure Tdm_Image.FlipImage(SrcDib : TDibGraphic;
                              DstDib : TDibGraphic;
                              FlipType : TFlipType);
var
  HorzTransform : TFlipHorizontalTransform;
  VertTransform : TFlipVerticalTransform;
begin
  if FlipType = ftHorizontal then
  begin
    HorzTransform := TFlipHorizontalTransform.Create;
    try
      HorzTransform.ApplyOnDest(SrcDib,DstDib);
    finally
      FreeAndNil(HorzTransform);
    end;
  end
  else
  begin
    VertTransform := TFlipVerticalTransform.Create;
    try
      VertTransform.ApplyOnDest(SrcDib,DstDib);
    finally
      FreeAndNil(VertTransform);
    end;
  end;
end;

procedure Tdm_Image.DoGraphicDeskew(SrcDib : TDibGraphic;
                                    DstDib : TDibGraphic);
var
  Transform : TNBDeskewTransform;
begin
  Transform := TNBDeskewTransform.Create;
  try
    Transform.ApplyOnDest(SrcDib, DstDib);
  finally
    FreeAndNil(Transform);
  end;
end;

procedure Tdm_Image.DoGraphicBorderClean(SrcDib : TDibGraphic;
                                         DstDib : TDibGraphic);
var
  Transform : TNBCleanupBorderTransform;
begin
  Transform := TNBCleanupBorderTransform.Create;
  try
    Transform.ApplyOnDest(SrcDib, DstDib);
  finally
    FreeAndNil(Transform);
  end;
end;

Procedure Tdm_Image.DoRotatDIB(TmpDib: TDibGraphic;
                               Angle: Double);
Var
  RotateTransform: TRotateTransform;
  Graphic: TDibGraphic;
begin
  Graphic := TDibGraphic.Create;
  RotateTransform := TRotateTransform.Create;
  Try
   Graphic.Assign(TmpDib);
   RotateTransform.Angle := Angle;
   RotateTransform.ApplyOnDest(Graphic, TmpDib);
  Finally
   FreeAndNil(Graphic);
   FreeAndNil(RotateTransform);
  End;
end;

procedure Tdm_Image.RotateIfNeed(FileName: String;
                                 var Pic: TDibGraphic;
                                 Portret: Boolean);
begin
  Pic.Clear;
  if LoadImageToDib(FileName, Pic, true) then
   if (Portret = (Pic.Width > Pic.Height)) and (Pic.Width <> Pic.Height) then
    dm_Image.DoRotatDIB(Pic, 90);
end;

Procedure Tdm_Image.DoInvertDIB(TmpDib: TDibGraphic);
var
  Transform: TNegativeTransform;
  FUndoGraphic: TDibGraphic;
begin
  FUndoGraphic := TDibGraphic.Create;
  Try
    FUndoGraphic.Assign(TmpDib);
    Transform := TNegativeTransform.Create;
    try
      Transform.ApplyOnDest(FUndoGraphic, TmpDib);
    finally
      FreeAndNil(Transform);
    end;
  Finally
   FreeAndNil(FUndoGraphic);
  End;
end;

function Tdm_Image.IsFileBMP(FileName: String): Boolean;
Var
  ExtName: String;
begin
  ExtName := J_ExtFileName(FileName);
  Result := False;
  IF UpperCase(ExtName) = UpperCase('BMP') Then
   Result := true;
end;

function Tdm_Image.IsFileJpeg(FileName: String): Boolean;
Var
  ExtName: String;
begin
  ExtName := J_ExtFileName(FileName);
  Result := False;
  IF (UpperCase(ExtName) = UpperCase(JPG_EXTENTION_NAME[1])) OR
    (UpperCase(ExtName) = UpperCase(JPG_EXTENTION_NAME[2])) Then
   Result := true;
end;

function Tdm_Image.IsFileTiff(FileName: String): Boolean;
Var
  ExtName: String;
begin
  ExtName := J_ExtFileName(FileName);
  Result := true;
  IF (UpperCase(ExtName) <> UpperCase(TIF_EXTENTION_NAME[1])) And
    (UpperCase(ExtName) <> UpperCase(TIF_EXTENTION_NAME[2])) Then
   Result := False;
end;

function  Tdm_Image.IsFilePNG(FileName: String): Boolean;
Var
  ExtName: String;
begin
  ExtName := J_ExtFileName(FileName);
  Result := (UpperCase(ExtName) = UpperCase('Png'));
end;

function  Tdm_Image.IsFilePCX(FileName: String): Boolean;
Var
  ExtName: String;
begin
  ExtName := J_ExtFileName(FileName);
  Result := (UpperCase(ExtName) = UpperCase('Pcx'));
end;

function  Tdm_Image.IsFileWMF(FileName: String): Boolean;
Var
  ExtName: String;
begin
  ExtName := J_ExtFileName(FileName);
  Result := (UpperCase(ExtName) = UpperCase('Wmf'));
end;

function Tdm_Image.IsFileDCX(FileName: String): Boolean;
Var
  ExtName: String;
begin
  ExtName := J_ExtFileName(FileName);
  Result := (UpperCase(ExtName) = UpperCase('Dcx'));
end;

function Tdm_Image.IsFilePDF(FileName: String): Boolean;
Var
  ExtName: String;
begin
  ExtName := J_ExtFileName(FileName);
  Result := (UpperCase(ExtName) = UpperCase('PDF'));
end;

function Tdm_Image.IsFileDWG(FileName: String): Boolean;
Var
  ExtName: String;
begin
  ExtName := J_ExtFileName(FileName);
  Result := (UpperCase(ExtName) = UpperCase('DWG'));
end;

function Tdm_Image.IsFileHTML(FileName: String): Boolean;
begin
  Result := IsImageShowHTMLViewer(FileName)
end;

function Tdm_Image.IsFileDOC(FileName: String): Boolean;
Var
  ExtName: String;
begin
  ExtName := J_ExtFileName(FileName);
  Result := (UpperCase(ExtName) = UpperCase('DOC'));
end;

function Tdm_Image.IsFileDOCX(FileName: String): Boolean;
Var
  ExtName: String;
begin
  ExtName := J_ExtFileName(FileName);
  Result := (UpperCase(ExtName) = UpperCase('DOCX'));
end;

function  Tdm_Image.IsFileWord(FileName : WideString) : Boolean;
Const
  IMAGE_WORD_VEIWER : Array [1..2] of String = ('DOC',
                                                'DOCX');
Var
  ExtName : WideString;
  CurImg  : Integer;
begin
  ExtName := J_ExtFileName(FileName);
  Result := False;
  For CurImg := Low(IMAGE_WORD_VEIWER) To High(IMAGE_WORD_VEIWER) Do
  begin
    IF (UpperCase(IMAGE_WORD_VEIWER[CurImg]) = UpperCase(ExtName)) Then
    begin
      Result := True;
      Break;
    end;
  end;
end;

function  Tdm_Image.IsFileXLS(FileName: String): Boolean;
Var
  ExtName: String;
begin
  ExtName := J_ExtFileName(FileName);
  Result := (UpperCase(ExtName) = UpperCase('XLS'));
end;

function  Tdm_Image.IsFileXLSX(FileName: String): Boolean;
Var
  ExtName: String;
begin
  ExtName := J_ExtFileName(FileName);
  Result := (UpperCase(ExtName) = UpperCase('XLSX'));
end;

function  Tdm_Image.IsFileExcel(FileName : WideString) : Boolean;
Const
  IMAGE_XLS_VEIWER : Array [1..2] of String = ('XLS',
                                               'XLSX');
Var
  ExtName : WideString;
  CurImg  : Integer;
begin
  ExtName := J_ExtFileName(FileName);
  Result := False;
  For CurImg := Low(IMAGE_XLS_VEIWER) To High(IMAGE_XLS_VEIWER) Do
  begin
    IF (UpperCase(IMAGE_XLS_VEIWER[CurImg]) = UpperCase(ExtName)) Then
    begin
      Result := True;
      Break;
    end;
  end;
end;

function  Tdm_Image.IsFilePPT(FileName: String): Boolean;
Var
  ExtName: String;
begin
  ExtName := J_ExtFileName(FileName);
  Result := (UpperCase(ExtName) = UpperCase('PPT'));
end;

function  Tdm_Image.IsFilePPTX(FileName: String): Boolean;
Var
  ExtName: String;
begin
  ExtName := J_ExtFileName(FileName);
  Result := (UpperCase(ExtName) = UpperCase('PPTX'));
end;

function  Tdm_Image.IsFilePowerPoint(FileName : WideString) : Boolean;
Const
  IMAGE_PPT_VEIWER : Array [1..2] of String = ('PPT',
                                               'PPTX');
Var
  ExtName : WideString;
  CurImg  : Integer;
begin
  ExtName := J_ExtFileName(FileName);
  Result := False;
  For CurImg := Low(IMAGE_PPT_VEIWER) To High(IMAGE_PPT_VEIWER) Do
  begin
    IF (UpperCase(IMAGE_PPT_VEIWER[CurImg]) = UpperCase(ExtName)) Then
    begin
      Result := True;
      Break;
    end;
  end;
end;

function  Tdm_Image.IsFileOutLookMsg(FileName: String): Boolean;
Var
  ExtName: String;
begin
  ExtName := J_ExtFileName(FileName);
  Result := (UpperCase(ExtName) = UpperCase('MSG'));
end;

function Tdm_Image.IsImageShowHTMLViewer(FileName: String): Boolean;
Var
  ExtName: String;
  CurImg: Integer;
begin
  Result := False;
  ExtName := J_ExtFileName(FileName);
  For CurImg := Low(IMAGE_HTML_VIEWER) To High(IMAGE_HTML_VIEWER) Do
  begin
    IF (UpperCase(IMAGE_HTML_VIEWER[CurImg]) = UpperCase(ExtName)) Then
    begin
      Result := true;
      Break;
    end;
  end;
end;

function Tdm_Image.IsImageShowPDFViewer(FileName: String): Boolean;
Var
  ExtName: String;
begin
  Result := False;
  ExtName := J_ExtFileName(FileName);
  IF UpperCase(ExtName) = UpperCase('PDF') Then
    Result := true;
end;

function Tdm_Image.IsImageShowDWGViewer(FileName: String): Boolean;
Var
  ExtName: String;
begin
  Result := False;
  ExtName := J_ExtFileName(FileName);
  IF UpperCase(ExtName) = UpperCase('DWG') Then
    Result := true;
end;

function Tdm_Image.IsImageCanAddedToPDF(FileName: String): Boolean;
Var
  ExtName: String;
  CurImg: Integer;
begin
  Result := False;
  ExtName := J_ExtFileName(FileName);
  ExtName := Trim(ExtName);
  For CurImg := Low(IMAGE_ADD_TO_PDF) To High(IMAGE_ADD_TO_PDF) Do
  begin
    IF (UpperCase(IMAGE_ADD_TO_PDF[CurImg]) = UpperCase(ExtName)) Then
    begin
      Result := true;
      Break;
    end;
  end;
end;

function Tdm_Image.IsImageShowenInternal(FileName: String): Boolean;
Var
  ExtName: String;
  CurImg: Integer;
begin
  Result := False;
  ExtName := J_ExtFileName(FileName);
  ExtName := Trim(ExtName);
  For CurImg := Low(IMAGE_EXT_INTERNAL_SHOW) To High(IMAGE_EXT_INTERNAL_SHOW) Do
  begin
    IF (UpperCase(IMAGE_EXT_INTERNAL_SHOW[CurImg]) = UpperCase(ExtName)) Then
    begin
      Result := true;
      Break;
    end;
  end;
end;

function Tdm_Image.IsImageCanLoadInScan(FileName: String): Boolean;
Var
  ExtName: String;
  CurImg: Integer;
begin
  Result := False;
  ExtName := J_ExtFileName(FileName);
  ExtName := Trim(ExtName);
  For CurImg := Low(IMAGE_SCANLOAD_SUPPORTED) To High(IMAGE_SCANLOAD_SUPPORTED) Do
  begin
    IF (UpperCase(IMAGE_SCANLOAD_SUPPORTED[CurImg]) = UpperCase(ExtName)) Then
    begin
      Result := true;
      Break;
    end;
  end;
end;

function Tdm_Image.IsImageDragAndDropSupported(FileName: String): Boolean;
Var
  ExtName: String;
  CurImg: Integer;
begin
  Result := False;
  ExtName := J_ExtFileName(FileName);
  ExtName := Trim(ExtName);
  For CurImg := Low(IMAGE_DRAGDROP_SUPPORTED) To High(IMAGE_DRAGDROP_SUPPORTED) Do
  begin
    IF (UpperCase(IMAGE_DRAGDROP_SUPPORTED[CurImg]) = UpperCase(ExtName)) Then
    begin
      Result := true;
      Break;
    end;
  end;
end;

function Tdm_Image.IsItOfficeVersion(FileName: String): Boolean;
Var
  ExtName: String;
  CurImg: Integer;
begin
  Result := False;
  ExtName := J_ExtFileName(FileName);
  ExtName := Trim(ExtName);
  For CurImg := Low(IMAGE_OFFICE_VERSION) To High(IMAGE_OFFICE_VERSION) Do
  begin
    IF (UpperCase(IMAGE_OFFICE_VERSION[CurImg]) = UpperCase(ExtName)) Then
    begin
      Result := true;
      Break;
    end;
  end;
end;

function Tdm_Image.IsFileCanBePrinted(FileName: String): Boolean;
Var
  ExtName: String;
  CurImg: Integer;
begin
  Result := False;
  ExtName := J_ExtFileName(FileName);
  ExtName := Trim(ExtName);
  For CurImg := Low(IMAGE_SEND_TO_PRINTER) To High(IMAGE_SEND_TO_PRINTER) Do
  begin
    IF (UpperCase(IMAGE_SEND_TO_PRINTER[CurImg]) = UpperCase(ExtName)) Then
    begin
      Result := true;
      Break;
    end;
  end;
end;

function Tdm_Image.IsFileCanBeGraphicPrinted(FileName: String): Boolean;
Var
  ExtName: String;
  CurImg: Integer;
begin
  Result := False;
  ExtName := J_ExtFileName(FileName);
  ExtName := Trim(ExtName);
  For CurImg := Low(IMAGE_SEND_TO_GRAPHICPRINTER) To High(IMAGE_SEND_TO_GRAPHICPRINTER) Do
  begin
    IF (UpperCase(IMAGE_SEND_TO_GRAPHICPRINTER[CurImg]) = UpperCase(ExtName)) Then
    begin
      Result := true;
      Break;
    end;
  end;
end;

function Tdm_Image.ImageCanViewByPreView(FileName: String): Boolean;
Var
  ExtnFile: String;
begin
  ExtnFile := J_ExtFileName(FileName);
  Result := (UpperCase(ExtnFile) = UpperCase('Pdf')) Or
            (UpperCase(ExtnFile) = UpperCase('Doc')) Or
            (UpperCase(ExtnFile) = UpperCase('Docx')) Or
            (UpperCase(ExtnFile) = UpperCase('Htm')) Or
            (UpperCase(ExtnFile) = UpperCase('Html')) Or
            (UpperCase(ExtnFile) = UpperCase('Xml')) Or
            (UpperCase(ExtnFile) = UpperCase('Xls')) Or
            (UpperCase(ExtnFile) = UpperCase('Tif')) Or
            (UpperCase(ExtnFile) = UpperCase('Tiff')) Or
            (UpperCase(ExtnFile) = UpperCase('Jpg')) Or
            (UpperCase(ExtnFile) = UpperCase('Jpeg')) Or
            (UpperCase(ExtnFile) = UpperCase('Bmp')) Or
            (UpperCase(ExtnFile) = UpperCase('Png')) Or
            (UpperCase(ExtnFile) = UpperCase('Ico')) Or
            (UpperCase(ExtnFile) = UpperCase('emf')) Or
            (UpperCase(ExtnFile) = UpperCase('wmf'));
end;

function Tdm_Image.ImageCanViewByHtmlViewer(FileName: String): Boolean;
Var
  ExtnFile: String;
begin
  ExtnFile := J_ExtFileName(FileName);
  Result := (UpperCase(ExtnFile) = UpperCase('Doc')) Or
            (UpperCase(ExtnFile) = UpperCase('Docx')) Or
            (UpperCase(ExtnFile) = UpperCase('Htm')) Or
            (UpperCase(ExtnFile) = UpperCase('Html')) Or
            (UpperCase(ExtnFile) = UpperCase('Xml')) Or
            (UpperCase(ExtnFile) = UpperCase('Xls'))Or
            (UpperCase(ExtnFile) = UpperCase('Xlsx'));
end;

function Tdm_Image.CanAddAnnonation(FileName: String): Boolean;
Var
  CurImg: Integer;
  ExtName: String;
begin
  Result := False;
  ExtName := J_ExtFileName(FileName);
  ExtName := Trim(ExtName);
  For CurImg := Low(CAN_ADD_ANNONATION_TO_IMAGE) To High
    (CAN_ADD_ANNONATION_TO_IMAGE) Do
  begin
    IF (UpperCase(CAN_ADD_ANNONATION_TO_IMAGE[CurImg]) = UpperCase(ExtName)) Then
    begin
      Result := true;
    end;
  end;
end;

function Tdm_Image.GetImagePageCount(FileName: String): Integer;
Var
  ExtName: String;
  CurImg: Integer;
  FHndl: Integer;
  Rslt: Integer;
  UnlockResult: Integer;
  PDFLibrary: TQuickPDF;
  SrcStream: TFileStream;
  SrcDcx: TDcxGraphic;
  SrcTif: TTiffGraphic;
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

          Rslt := 0;
          FHndl := PDFLibrary.DAOpenFileReadOnly(FileName, '');
          IF FHndl > 0 Then
          begin
            Rslt := PDFLibrary.DAGetPageCount(FHndl);
          end;

          IF Rslt > 0 Then
            Result := Rslt;
          PDFLibrary.DACloseFile(FHndl);
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
              FreeAndNil(SrcTif);
            Except;
            End;
          End;
        Finally
          Try
            FreeAndNil(SrcStream);
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
              FreeAndNil(SrcDcx);
            Except;
            End;
          End;
        Finally
          Try
            FreeAndNil(SrcStream);
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

function Tdm_Image.GetTifFrameCount(FileName: String): Integer;
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
      Try
        FreeAndNil(SrcTif);
      Except;
      End;
    End;
  Finally
    Try
      FreeAndNil(SrcStream);
    Except;
    End;
    Try
      SrcStream := nil;
    Except;
    End;
  End;
end;

function Tdm_Image.CanConvertTo_PDF(FileName: String): Boolean;
Var
  ExtName: String;
  CurImg: Integer;
begin
  Result := False;
  ExtName := J_ExtFileName(FileName);
  For CurImg := Low(CONVERT_TO_PDF) To High(CONVERT_TO_PDF) Do
  begin
    IF (UpperCase(CONVERT_TO_PDF[CurImg]) = UpperCase(ExtName)) Then
    begin
      Result := true;
      Break;
    end;
  end;
end;

function Tdm_Image.CanConvertTo_TIF(FileName: String): Boolean;
Var
  ExtName: String;
  CurImg: Integer;
begin
  Result := False;
  ExtName := J_ExtFileName(FileName);
  For CurImg := Low(CONVERT_TO_TIF) To High(CONVERT_TO_TIF) Do
  begin
    IF (UpperCase(CONVERT_TO_TIF[CurImg]) = UpperCase(ExtName)) Then
    begin
      Result := true;
      Break;
    end;
  end;
end;

function Tdm_Image.CanConvertTo_JPG(FileName: String): Boolean;
Var
  ExtName: String;
  CurImg: Integer;
begin
  Result := False;
  ExtName := J_ExtFileName(FileName);
  For CurImg := Low(CONVERT_TO_JPG) To High(CONVERT_TO_JPG) Do
  begin
    IF (UpperCase(CONVERT_TO_JPG[CurImg]) = UpperCase(ExtName)) Then
    begin
      Result := true;
      Break;
    end;
  end;
end;

function Tdm_Image.CanUpdateFileSize(FileName: String): Boolean;
Var
  ExtName: String;
  CurImg: Integer;
begin
  Result := False;
  ExtName := J_ExtFileName(FileName);
  For CurImg := Low(IMAGE_CAN_GET_PAGE_COUNT) To High(IMAGE_CAN_GET_PAGE_COUNT) Do
  begin
    IF (UpperCase(IMAGE_CAN_GET_PAGE_COUNT[CurImg]) = UpperCase(ExtName)) Then
    begin
      Result := true;
      Break;
    end;
  end;
end;

function Tdm_Image.CanScanMoreToDocument(FileName: String): Boolean;
Var
  ExtName: String;
  CurImg: Integer;
begin
  Result := False;
  ExtName := J_ExtFileName(FileName);
  For CurImg := Low(CAN_SCAN_MORE_TO_DOCUMENT) To High
    (CAN_SCAN_MORE_TO_DOCUMENT) Do
  begin
    IF (UpperCase(CAN_SCAN_MORE_TO_DOCUMENT[CurImg]) = UpperCase(ExtName)) Then
    begin
      Result := true;
      Break;
    end;
  end;
end;

function Tdm_Image.LoadImageToTif(FileName: String;
                                  Var Tmp: TTiffGraphic;
                                  SilentMode: Boolean = False;
                                  LoadJustFirstPage: Boolean = False;
                                  ConvertToBW : Boolean = True): Boolean;
Var
  SrcRect: TRect;
  DstRect: TRect;
  TmpBMP : TBitmapGraphic;
  TmpICO : TIconGraphic;
  TmpJPG : TJpegGraphic;
  TmpPCX : TPcxGraphic;
  TmpWmf : TMetaFileGraphic;
  TmpDcx : TDcxGraphic;
  TmpPNG : TPngGraphic;
  TmpTGA : TTgaGraphic;
  TmpDib : TDibGraphic;
  TmpGif : TImage;
  ExtFileName : String;
  TmpFileName : WideString;
begin
  Result := False;
  Tmp.ClearAll;

  IF Not FileExists(FileName) Then
  begin
    IF SilentMode Then
    begin
      Tmp.Assign(ImageNotFoundBmp);
      Result := true;
      Exit;
    end
    else
      Raise EAppError.Create(Format(dm_SiLang.siLangLinked.GetTextOrDefault('IDS_26' (* 'קובץ לא נמצא' *) ), [FileName]));
  end;

  Tmp.MultiLoad := False;

  ExtFileName := J_ExtFileName(FileName);

  Try
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
    If UpperCase(ExtFileName) = UpperCase('PDF') Then
    begin
      if not LoadJustFirstPage then
      begin
        if not uPdfUtil.RenderPDFToMultiTiff(FileName,
                                             Tmp,
                                             ParamArea.Param_ConvertToPDF_DPI,
                                             ErrorCode) Then
        begin
          Tmp.ClearAll;
        end;
      end
      else
      begin
        TmpDib := TDibGraphic.Create;
        Try
          if RenderFirstPageToDib(FileName,TmpDib) Then
          begin
            Tmp.Assign(TmpDib);
          end
          else
          begin
            Tmp.ClearAll;
          end;
        Finally
          FreeAndNil(TmpDib);
        End;
      end;

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
    IF IsFileTiff(FileName) Then // 'TIF'..'TIFF'
    begin
      Tmp.MultiLoad := true;
      if LoadJustFirstPage then
        Tmp.MultiLoad := False;
      Try
        Tmp.LoadFromFile(FileName);
      Except;
        TmpJPG := TJpegGraphic.Create;
        Try
          TmpJPG.LoadFromFile(FileName);
          Tmp.Assign(TmpJPG);
        Finally
          FreeAndNil(TmpJPG);
        End;
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
      Except
        on e : Exception Do
        begin
          dm_Main.MzWriteLog(e.Message, MzLog_FromProcess);
          Tmp.ClearAll;
        end;
      End;
    end
    else
    begin
      Try
        TmpDib := TDibGraphic.Create;
        Try
          TmpDib.LoadFromFile(FileName);
          Tmp.Assign(TmpDib);
        Finally
          FreeAndNil(TmpDib);
        End;
      Except
        On e: Exception Do
        begin
          dm_Main.MzWriteLog(e.Message,MzLog_FromProcess);
          Tmp.ClearAll;
        end;
      end;
    end;
    Result := Not Tmp.IsEmpty;
  Except;
    Result := False;
  End;
end;

function Tdm_Image.LoadImageToDib(FileName: String;
                                  Var Tmp: TDibGraphic;
                                  SilentMode: Boolean = False): Boolean;
Var
  TmpBMP  : TBitmapGraphic;
  TmpICO  : TIconGraphic;
  TmpJPG  : TJpegGraphic;
  TmpWmf  : TMetaFileGraphic;
  TmpPCX  : TPcxGraphic;
  TmpPNG  : TPngGraphic;
  TmpTGA  : TTgaGraphic;
  TmpTif  : TTiffGraphic;
  TmpDib  : TDibGraphic;
  TmpDCX  : TDcxGraphic;
  TmpGif  : TImage;

  SrcRect : TRect;
  DstRect : TRect;

  ExtFileName : String;
  SomeIcon    : TIcon;
  TmpFileName : String;
begin
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
      Raise EAppError.Create(Format(dm_SiLang.siLangLinked.GetTextOrDefault('IDS_26' (* 'קובץ לא נמצא' *) ), [FileName]));
  end;

  IF Not FileExists(FileName) Then
  begin
    IF SilentMode Then
    begin
      Result := true;
      Exit;
    end
    else
      Raise EAppError.Create(Format(dm_SiLang.siLangLinked.GetTextOrDefault('IDS_26' (* 'קובץ לא נמצא' *) ), [FileName]));
  end;

  ExtFileName := ExtractFileExt(FileName);
  IF (copy(ExtFileName, 1, 1) = '.') Then
    ExtFileName := copy(ExtFileName, 2, Length(ExtFileName) - 1);

  If UpperCase(ExtFileName) <> UpperCase('PDF') Then
  begin
    if not IsImageShowenInternal(FileName) then
    begin
      SomeIcon := TIcon.Create;
      Try
        if GetIconFromFile(FileName, SomeIcon) then
          Tmp.Assign(SomeIcon)
        else
          Tmp.Assign(ImageNotFoundBmp);
      Finally
        FreeAndNil(SomeIcon);
      End;
      Result := true;
      Exit;
    end;
  end;

  IF Not FileExists(FileName) Then
   Exit;

  Try
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
    If UpperCase(ExtFileName) = UpperCase('PDF') Then
    begin
      if not uPdfUtil.RenderFirstPageToDib(FileName,
                                           Tmp,
                                           ParamArea.Param_ConvertToPDF_DPI,
                                           ErrorCode) Then
        Tmp.Clear;
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
    If (UpperCase(ExtFileName) = UpperCase('TIF')) Or
       (UpperCase(ExtFileName) = UpperCase('TIFF')) Then
    begin
      TmpTif := TTiffGraphic.Create;
      TmpTif.MultiLoad := False;
      Tmp.MultiLoad := False;
      Try
        TmpTif.LoadFromFile(FileName);
        Tmp.Assign(TmpTif);
      Finally
        FreeAndNil(TmpTif);
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
    end
    else
    begin
      Try
        TmpDib := TDibGraphic.Create;
        Try
          TmpDib.LoadFromFile(FileName);
          Tmp.Assign(TmpDib);
        Finally
          FreeAndNil(TmpDib);
        End;
      Except;
        Tmp.Clear;
      end;
    end;

    Result := Not Tmp.IsEmpty;
  Except;
    Tmp.Clear;
    Result := False;
  End;
end;

function  Tdm_Image.LoadImageToEnvision(Image: TImage;
                                         Var Tmp: TJpegGraphic): Boolean;
Var
  Ms  : TMemoryStream;
begin
  Result := False;
  Try
    Ms  := TMemoryStream.Create;
    Try
      Image.Picture.Graphic.SaveToStream(Ms);
      Ms.Position := 0;
      Tmp.LoadFromStream(Ms);
    Finally
      FreeAndNil(Ms);
    End;
    Result := True;
  Except;
    Result := False;
  End;
end;

function  Tdm_Image.LoadPdfToDib(FileName: String;
                                 Var Tmp: TDibGraphic;
                                 SilentMode: Boolean = False): Boolean;
Var
  PdfFileName  : WideString;
  PDFLibrary   : TQuickPDF;
  MemoryStream : TMemoryStream;
  Rslt         : Integer;
  FHndl        : Integer;
  PageRef      : Integer;
  UnlockResult : Integer;
  TmpJpg       : TJPegGraphic;
  CurCursor    : TCursor;
begin
  Result := False;

  Tmp.Clear;
  PdfFileName := Trim(FileName);
  if Not FileExists(PdfFileName) then
    Exit;

  Try
    if not uPdfUtil.RenderFirstPageToDib(FileName,
                                         Tmp,
                                         ParamArea.Param_ConvertToPDF_DPI,
                                         ErrorCode) Then
      Tmp.Clear;
  Except
    on e:exception do
    begin
      Tmp.Clear;
      JustWriteToLog(e.Message);
      if not SilentMode then
        Raise;
    end;
  End;

  Result := Not Tmp.IsEmpty;
end;

function Tdm_Image.BreakTifToPages(FileName: String;
                                   ToDir: String;
                                   FileLists: TStringList): Boolean;
Var
  PageCount   : Integer;
  CurFrame    : Integer;
  SaveTif     : TTiffGraphic;
  NewFileName : String;
  MsgStr      : String;
begin
  Result := False;
  IF Not FileExists(FileName) Then
   Exit;
  IF Not IsFileTiff(FileName) Then
   Exit;
  Try
    FileLists.Clear;

    PageCount := GetTifFrameCount(FileName);
    IF PageCount <= 1 Then
    begin
      NewFileName := RemoveBackSlashChar(ToDir) + '\' + j_FileName(FileName);
      // ToDir is of the function param
      IF Not DirectoryExists(ToDir) Then
        ForceDirectories(ToDir);
      IF Not DirectoryExists(ToDir) Then
      begin
        MsgStr := Format(dm_SiLang.siLangLinked.GetTextOrDefault('IDS_27' (* 'כשלון ביצירת ספריה זמנית לפירוק קובץ TIF' *) ), [ToDir]);
        dm_Main.MzWriteLog(Trim(MsgStr),MzLog_FromProcess);
        Exit;
      end;

      Windows.CopyFile(PWideChar(FileName), PWideChar(NewFileName), False);
      IF FileExists(NewFileName) Then
      begin
        FileLists.Add(NewFileName);
        Result := true;
      end;
    end
    else
    begin
      Result := true;
      SaveTif := TTiffGraphic.Create;
      Try
        For CurFrame := 1 To PageCount Do
        begin
          NewFileName := RemoveBackSlashChar(ToDir) + '\' + J_FirstFileName
            (FileName) + '_' + IntToStr(CurFrame) + '.Tif';

          SaveTif.MultiLoad := False;
          SaveTif.SingleLoadFromFile(FileName, CurFrame);
          IF SaveTifToFile(SaveTif, NewFileName) Then
          begin
            FileLists.Add(NewFileName);
          end
          else
          begin
            Result := False;
            Break;
          end;
        end;
      Finally
       FreeAndNil(SaveTif);
      End;
    end;
  Finally
  End;
end;

function  Tdm_Image.BreakPdfToTiffPages(PdfFilename : String;
                                        ToDir       : String;
                                        Dpi         : Integer;
                                        FileLists   : TStringList) : Boolean;
Var
  Rslt          : Integer;
  ToFileName    : String;
  CheckFileName : String;
  UnlockResult  : Integer;
  PDFLibrary    : TQuickPDF;
  CurFile       : Integer;
  CurCursor     : TCursor;
begin
  Result := False;
  FileLists.Clear;

  ForceDirectories(ToDir);
  IF Not DirectoryExists(ToDir) Then
    Beep;

  ToFileName := RemoveBackSlashChar(ToDir) + '\Tmp%p.Tif';
  PDFLibrary := TQuickPDF.Create;
  Try
    UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
    IF UnlockResult <> 1 then
      Exit;

    IF Not FileExists(PdfFileName) Then
      Exit;

    CurCursor := Screen.Cursor;
    Try
      Screen.Cursor := crHourGlass;
      PDFLibrary.LoadFromFile(PdfFileName,'');
      Rslt := PDFLibrary.RenderDocumentToFile(DPI,1,PDFLibrary.PageCount,7{tif},ToFileName);
      IF Rslt = 1 Then
      begin
        For CurFile := 1 To 10000 Do
        begin
          CheckFileName := RemoveBackSlashChar(ToDir) + '\Tmp' + IntToStr(CurFile) + '.Tif';
          IF Not FileExists(CheckFileName) Then
            Break;
          FileLists.Add(CheckFileName);
        end;
      end;
    Finally
      Screen.Cursor := CurCursor;
    End;
  Finally
    FreeAndNil(PDFLibrary);
  End;

  Result := (FileLists.Count > 0);
end;

procedure Tdm_Image.ShowEnvisionInternal(ImageRec: TImageRec);
Var
  CurCursor : TCursor;
begin
  IF Not FileExists(ImageRec.FileName) Then
    Exit;
  IF Not IsImageShowenInternal(ImageRec.FileName) Then
    Exit;

  EnShowImageFrm := TEnShowImageFrm.Create(Application);
  Try
    CurCursor := Screen.Cursor;
    Screen.Cursor := crHourGlass;
    EnShowImageFrm.Name := FormName_EN + IntToStr(GetTickCount);
    EnShowImageFrm.StayOnTopMode := UserDataRec.User_ViewImgFormStayOnTop;
    EnShowImageFrm.ImageRec := ImageRec;
    EnShowImageFrm.Show;
  Finally
    Screen.Cursor := CurCursor;
  End;
end;

procedure Tdm_Image.ShowGeneralInternal(ImageRec: TImageRec);
Var
  CurCursor : TCursor;
  UserViewPdfKind : Integer;
begin

  IF Not FileExists(ImageRec.FileName) Then
    Exit;

  GeneralViewerFrm := TGeneralViewerFrm.Create(Application);
  Try
    CurCursor := Screen.Cursor;
    Screen.Cursor := crHourGlass;
    GeneralViewerFrm.Name := FormName_GENERAL + IntToStr(GetTickCount);
    GeneralViewerFrm.MzIdent         := ImageRec.Ident;
    GeneralViewerFrm.GeneralFileName := ImageRec.FileName;
    GeneralViewerFrm.Show;
    GeneralViewerFrm.BringToFront;
  Finally
    Screen.Cursor := CurCursor;
  End;
end;

procedure Tdm_Image.ShowOfficeInternal(ImageRec: TImageRec);
Var
  CurCursor : TCursor;
  UserViewPdfKind : Integer;
begin

  IF Not FileExists(ImageRec.FileName) Then
    Exit;
  IF Not IsFileOffice(ImageRec.FileName) Then
    Exit;

  OfficeViewerFrm := TOfficeViewerFrm.Create(Application);
  Try
    CurCursor := Screen.Cursor;
    Screen.Cursor := crHourGlass;
    OfficeViewerFrm.Name := FormName_OFFICE + IntToStr(GetTickCount);
    OfficeViewerFrm.MzIdent        := ImageRec.Ident;
    OfficeViewerFrm.OfficeFileName := ImageRec.FileName;
    OfficeViewerFrm.Show;
    OfficeViewerFrm.BringToFront;
  Finally
    Screen.Cursor := CurCursor;
  End;
end;

procedure Tdm_Image.ShowOutlookInternal(ImageRec: TImageRec);
Var
  CurCursor : TCursor;
  UserViewPdfKind : Integer;
begin
  IF Not FileExists(ImageRec.FileName) Then
    Exit;
  IF Not IsFileOutLookMsg(ImageRec.FileName) Then
    Exit;

  OutlookViewerFrm := TOutlookViewerFrm.Create(Application);
  Try
    CurCursor := Screen.Cursor;
    Screen.Cursor := crHourGlass;
    OutlookViewerFrm.Name := FormName_OFFICE + IntToStr(GetTickCount);
    OutlookViewerFrm.MzIdent         := ImageRec.Ident;
    OutlookViewerFrm.OutlookFileName := ImageRec.FileName;
    OutlookViewerFrm.Show;
    OutlookViewerFrm.BringToFront;
  Finally
    Screen.Cursor := CurCursor;
  End;
end;

procedure Tdm_Image.ShowPDFInternal(ImageRec: TImageRec);
Var
  CurCursor : TCursor;
  UserViewPdfKind : Integer;
begin

  IF Not FileExists(ImageRec.FileName) Then
    Exit;
  IF Not IsFilePDF(ImageRec.FileName) Then
    Exit;

  UserViewPdfKind := dm_Ro.GetViewPdfMethode;
  If (UserViewPdfKind = PDF_ViewInternal) or ImageRec.JustInternal Then
  begin
    PDFViewerFrm := TPDFViewerFrm.Create(Application);
    PDFViewerFrm.Name := FormName_PDF + IntToStr(GetTickCount);
    CurCursor := Screen.Cursor;
    Try
     Screen.Cursor := crHourGlass;
     PDFViewerFrm.PdfFileName := ImageRec.FileName;
     PDFViewerFrm.MzIdent     := ImageRec.Ident;
     PDFViewerFrm.Show;
     PDFViewerFrm.BringToFront;
    Finally
     Screen.Cursor := CurCursor;
    End;
  end
  else
  begin
    ShowImage_External(ImageRec);
  end;
end;

procedure Tdm_Image.ShowHtmlInternal(ImageRec: TImageRec;
                                      MailSubject: String);
Var
  CurCursor: TCursor;
begin
  HTMLViewerFrm := THTMLViewerFrm.Create(Application);
  Try
    CurCursor := Screen.Cursor;
    Try
      Screen.Cursor := crHourGlass;
      HTMLViewerFrm.Name := FormName_HTML + IntToStr(GetTickCount);
      HTMLViewerFrm.MailSubject  := MailSubject;
      HTMLViewerFrm.HtmlText     := '';
      HTMLViewerFrm.HtmlFileName := ImageRec.FileName;
    Finally
      Screen.Cursor := CurCursor;
    End;
    HTMLViewerFrm.Show;
  Finally
  End;
end;

procedure Tdm_Image.ShowImage_Internal(ImageRec: TImageRec);
begin
  IF IsImageShowHTMLViewer(ImageRec.FileName) Then
  begin
    IF ParamArea.Param_ViewHtmlInInternalViewer Then
    begin
      ShowHtmlInternal(ImageRec, ImageRec.FileName);
      Exit;
    End;
  end;

  IF IsImageShowPDFViewer(ImageRec.FileName) Then
  begin
    IF ParamArea.Param_ViewPdfInInternaViewer Or (ImageRec.JustInternal And (UserDataRec.User_Level >= 90)) Then
    begin
      ShowPDFInternal(ImageRec);
      Exit;
    end;
  end;

  If (IsFileWord(ImageRec.FileName) And UserDataRec.User_ViewDocXViewr) Then
  begin
    ShowOfficeInternal(ImageRec);
    Exit;
  end;

  If (IsFileExcel(ImageRec.FileName) And UserDataRec.User_ViewXlsViewr) Then
  begin
    ShowOfficeInternal(ImageRec);
    Exit;
  end;

  If (IsFilePowerPoint(ImageRec.FileName) And UserDataRec.User_ViewPPTViewr) Then
  begin
    ShowOfficeInternal(ImageRec);
    Exit;
  end;

  if (IsFileOutlookMsg(ImageRec.FileName) And UserDataRec.User_ViewOutLookMsgViewr) Then
  begin
    ShowOutlookInternal(ImageRec);
    Exit;
  end;

  IF IsImageShowenInternal(ImageRec.FileName) Then
  begin
    IF ParamArea.Param_ViewImageInInternalViewer Or (ImageRec.JustInternal And (UserDataRec.User_Level >= 90)) Then
    begin
      ShowEnvisionInternal(ImageRec);
      Exit;
    end;
  end;

  ShowImage_External(ImageRec);
end;

procedure Tdm_Image.ShowImage_Internal(ImageRec: TImageRec;
                                        DoPreview: Boolean);
Var
  CurCursor: TCursor;
begin
  IF Not DoPreview Then
  begin
    ShowImage_Internal(ImageRec);
    Exit;
  end;

  IF ImageCanViewByHtmlViewer(ImageRec.FileName) Then
  begin
    ShowHtmlInternal(ImageRec, ImageRec.FileName);
    Exit;
  end;

  IF IsImageShowPDFViewer(ImageRec.FileName) Then
  begin
    ShowPDFInternal(ImageRec);
    Exit;
  end;

  IF IsImageShowenInternal(ImageRec.FileName) Then
  begin
    EnShowImageFrm := TEnShowImageFrm.Create(Application);
    EnShowImageFrm.Name := 'ImageEn' + IntToStr(GetTickCount);
    EnShowImageFrm.StayOnTopMode := UserDataRec.User_ViewImgFormStayOnTop;

    CurCursor := Screen.Cursor;
    Try
     Screen.Cursor := crHourGlass;
     EnShowImageFrm.ImageRec := ImageRec;
    Finally
     Screen.Cursor := CurCursor;
    End;
    EnShowImageFrm.Show;
    Exit;
  end;
end;

procedure Tdm_Image.ShowImage_External(ImageRec: TImageRec);
var
  FileName: String;
begin
  FileName := ImageRec.FileName;
  IF FileExists(FileName) Then
  begin
    ShowImage_External(FileName);
  end;
end;

procedure Tdm_Image.ShowImage_External(FileName: String);
var
  RsltCode : Cardinal;
  Found: Boolean;
  MsgStr : String;
  CatalogRec : TCatalogRec;
begin

  IF FileExists(FileName) Then
  begin
    Found := CallProgramAssociation(Application.Handle, FileName);
    IF Found Then
    begin
      JustWriteToLog('File : ' + FileName + ' exists. Associat program found');
      if ParamArea.Param_WriteActionLog then
      begin
        CatalogRec := dm_Main.GetCatalogDetailByFile(FileName);
        ActionLogRec := dm_Main.SetActionLogFromCatalogRec(CatalogRec);
        ActionLogRec.Log_Kind       := AL_OpenFileInExternalViewer;
        ActionLogRec.Log_ActionText := dm_SiLang.siLangLinked.GetTextOrDefault('IDS_28' (* 'פתיחת קובץ באמצעות עורך חיצוני' *) );
        ActionLogRec.Log_SqlText    := dm_SiLang.siLangLinked.GetTextOrDefault('IDS_29' (* 'פתיחה באמצעות תוכנית מוגדרת במערכת ההפעלה' *) );
        dm_Main.AddActionLog(ActionLogRec);
      end;

      Exit;
    end;

    JustWriteToLog('File : ' + FileName + ' exists. fail open file with Associat program');
    if ParamArea.Param_WriteActionLog then
    begin
      CatalogRec := dm_Main.GetCatalogDetailByFile(FileName);
      ActionLogRec := dm_Main.SetActionLogFromCatalogRec(CatalogRec);
      ActionLogRec.Log_Kind       := AL_OpenFileInExternalViewer;
      ActionLogRec.Log_ActionText := dm_SiLang.siLangLinked.GetTextOrDefault('IDS_28' (* 'פתיחת קובץ באמצעות עורך חיצוני' *) );
      ActionLogRec.Log_SqlText    := dm_SiLang.siLangLinked.GetTextOrDefault('IDS_30' (* 'פתיחה באמצעות הפעלת פקודת Shell-Open' *) );
      dm_Main.AddActionLog(ActionLogRec);
    end;

    JustWriteToLog('File : ' + FileName + ' exists. try open with ShellExecute');
    RsltCode := ShellExecute(Application.Handle, 'open', PChar('RUNDLL32.EXE'),
      PChar('shell32.dll,OpenAs_RunDLL ' + FileName), nil, SW_SHOWNORMAL);

    JustWriteToLog('File : ' + FileName + ' exists. open with ShellExecute, Result Code = ' + IntToStr(RsltCode));

    MsgStr := 'No Program Association With The Extention .' +
      J_ExtFileName(FileName) + #13 +
      'LOOK FOR - HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.'
      + J_ExtFileName(FileName) + '[Value of - Application]';

    dm_Main.MzWriteLog(MsgStr,MzLog_FromProcess);
  end
  else
  begin
    JustWriteToLog('File : ' + FileName + ' not exists. can not open file');
  end;
end;

procedure Tdm_Image.RemoveDot(Src: TDibGraphic; Dst: TDibGraphic; UseAgresiveMode : Boolean = False; CleanLoop : Integer = 1);
var
  X   : Integer;
  Y   : Integer;
  A   : TRGB;
  B01 : TRGB;
  B02 : TRGB;
  B03 : TRGB;
  B04 : TRGB;
  B05 : TRGB;
  B06 : TRGB;
  B07 : TRGB;
  B08 : TRGB;
  B09 : TRGB;
  B10 : TRGB;
  B11 : TRGB;
  B12 : TRGB;
  B13 : TRGB;
  B14 : TRGB;
  B15 : TRGB;
  B16 : TRGB;
  CurLoop : Integer;
  Cnt     : Integer;
  TmpDib  : TDibGraphic;
  ConvertToWhite    : Boolean;
  GamaTransform     : TGammaTransform;
  GamaDib           : TDibGraphic;
  ContrastTransform : TContrastTransform;
  SharpenTransform  : TSharpenTransform;
  DibSrc            : TDibGraphic;
  DibDst            : TDibGraphic;
  Transform         : TDespeckleTransform;
  DotCleaned        : Boolean;
begin
  Transform := TDespeckleTransform.Create;
  try
    Transform.ApplyOnDest(Src, Dst);
    Src.Assign(Dst);
  finally
    FreeAndNil(Transform);
  end;

  If CleanLoop < 1 Then
    CleanLoop := 1;
  If CleanLoop > 4 Then
    CleanLoop := 4;

  TmpDib := TDibGraphic.Create;
  Try
    TmpDib.Assign(Src);
    for CurLoop := 1 to CleanLoop do
    begin
      Dst.Assign(TmpDib);
      for X := 1 to Dst.Width - 2 do
      begin
        for Y := 1 to Dst.Height - 2 do
        begin
          DotCleaned := False;
          Cnt := 0;
          A := Dst.Rgb[X, Y];
          if ((A.Red < 30) and (A.Blue < 30) and (A.Green < 30)) Or
             ((A.Red + A.Blue + A.Green) < 90) then
          begin
            B01 := Dst.Rgb[X-1, Y-1];
            B02 := Dst.Rgb[X-1, Y];
            B03 := Dst.Rgb[X-1, Y+1];

            B04 := Dst.Rgb[X, Y-1];
            B05 := Dst.Rgb[X, Y+1];

            B06 := Dst.Rgb[X-2, Y-1];
            B07 := Dst.Rgb[X-2, Y];
            B08 := Dst.Rgb[X-2, Y+1];

            if (((B01.Red > 220) and (B01.Blue > 220) and (B01.Green > 220)) Or
                ((B01.Red + B01.Blue + B01.Green) > 600)) Then
              Cnt := Cnt + 1;

            if (((B02.Red > 220) and (B02.Blue > 220) and (B02.Green > 220)) Or
                ((B02.Red + B02.Blue + B02.Green) > 600)) Then
              Cnt := Cnt + 1;

            if (((B03.Red > 220) and (B03.Blue > 220) and (B03.Green > 220)) Or
                ((B03.Red + B03.Blue + B03.Green) > 600)) Then
              Cnt := Cnt + 1;

            if (((B04.Red > 220) and (B04.Blue > 220) and (B04.Green > 220)) Or
                ((B04.Red + B04.Blue + B04.Green) > 600)) Then
              Cnt := Cnt + 1;

            if (((B05.Red > 220) and (B05.Blue > 220) and (B05.Green > 220)) Or
                ((B05.Red + B05.Blue + B05.Green) > 600)) Then
              Cnt := Cnt + 1;

            if (((B06.Red > 220) and (B06.Blue > 220) and (B06.Green > 220)) Or
                ((B06.Red + B06.Blue + B06.Green) > 600)) Then
              Cnt := Cnt + 1;

            if (((B07.Red > 220) and (B07.Blue > 220) and (B07.Green > 220)) Or
                ((B07.Red + B07.Blue + B07.Green) > 600)) Then
              Cnt := Cnt + 1;

            if (((B08.Red > 220) and (B08.Blue > 220) and (B08.Green > 220)) Or
                ((B08.Red + B08.Blue + B08.Green) > 600)) Then
              Cnt := Cnt + 1;

            if Cnt > 7 Then
            begin
              A.Red   := 255;
              A.Blue  := 255;
              A.Green := 255;
              Dst.Rgb[X, Y] := A;
              DotCleaned := True;
            end;

            if UseAgresiveMode And Not DotCleaned then
            begin
              if (X > 2) And (X < Dst.Width - 2) And
                 (Y > 2) And (Y < Dst.Height - 2) Then
              begin
                B01 := Dst.Rgb[X-2, Y-2];
                B02 := Dst.Rgb[X-1, Y-2];
                B03 := Dst.Rgb[X, Y-2];
                B04 := Dst.Rgb[X+1, Y-2];
                B05 := Dst.Rgb[X+2, Y-2];

                B06 := Dst.Rgb[X+2, Y-1];
                B07 := Dst.Rgb[X+2, Y];
                B08 := Dst.Rgb[X+2, Y+1];

                B09 := Dst.Rgb[X-2, Y-1];
                B10 := Dst.Rgb[X-2, Y];
                B11 := Dst.Rgb[X-2, Y+1];

                B12 := Dst.Rgb[X-2, Y+2];
                B13 := Dst.Rgb[X-1, Y+2];
                B14 := Dst.Rgb[X, Y+2];
                B15 := Dst.Rgb[X+1, Y+2];
                B16 := Dst.Rgb[X+2, Y+2];

                if (((B01.Red > 220) and (B01.Blue > 220) and (B01.Green > 220)) Or
                    ((B01.Red + B01.Blue + B01.Green) > 600)) Then
                  Cnt := Cnt + 1;

                if (((B02.Red > 220) and (B02.Blue > 220) and (B02.Green > 220)) Or
                    ((B02.Red + B02.Blue + B02.Green) > 600)) Then
                  Cnt := Cnt + 1;

                if (((B03.Red > 220) and (B03.Blue > 220) and (B03.Green > 220)) Or
                    ((B03.Red + B03.Blue + B03.Green) > 600)) Then
                  Cnt := Cnt + 1;

                if (((B04.Red > 220) and (B04.Blue > 220) and (B04.Green > 220)) Or
                    ((B04.Red + B04.Blue + B04.Green) > 600)) Then
                  Cnt := Cnt + 1;

                if (((B05.Red > 220) and (B05.Blue > 220) and (B05.Green > 220)) Or
                    ((B05.Red + B05.Blue + B05.Green) > 600)) Then
                  Cnt := Cnt + 1;

                if (((B06.Red > 220) and (B06.Blue > 220) and (B06.Green > 220)) Or
                    ((B06.Red + B06.Blue + B06.Green) > 600)) Then
                  Cnt := Cnt + 1;

                if (((B07.Red > 220) and (B07.Blue > 220) and (B07.Green > 220)) Or
                    ((B07.Red + B07.Blue + B07.Green) > 600)) Then
                  Cnt := Cnt + 1;

                if (((B08.Red > 220) and (B08.Blue > 220) and (B08.Green > 220)) Or
                    ((B08.Red + B08.Blue + B08.Green) > 600)) Then
                  Cnt := Cnt + 1;

                if (((B09.Red > 220) and (B09.Blue > 220) and (B09.Green > 220)) Or
                    ((B09.Red + B09.Blue + B09.Green) > 600)) Then
                  Cnt := Cnt + 1;

                if (((B10.Red > 220) and (B10.Blue > 220) and (B10.Green > 220)) Or
                    ((B10.Red + B10.Blue + B10.Green) > 600)) Then
                  Cnt := Cnt + 1;

                if (((B11.Red > 220) and (B11.Blue > 220) and (B11.Green > 220)) Or
                    ((B11.Red + B11.Blue + B11.Green) > 600)) Then
                  Cnt := Cnt + 1;

                if (((B12.Red > 220) and (B12.Blue > 220) and (B12.Green > 220)) Or
                    ((B12.Red + B12.Blue + B12.Green) > 600)) Then
                  Cnt := Cnt + 1;

                if (((B13.Red > 220) and (B13.Blue > 220) and (B13.Green > 220)) Or
                    ((B13.Red + B13.Blue + B13.Green) > 600)) Then
                  Cnt := Cnt + 1;

                if (((B14.Red > 220) and (B14.Blue > 220) and (B14.Green > 220)) Or
                    ((B14.Red + B14.Blue + B14.Green) > 600)) Then
                  Cnt := Cnt + 1;

                if (((B15.Red > 220) and (B15.Blue > 220) and (B15.Green > 220)) Or
                    ((B15.Red + B15.Blue + B15.Green) > 600)) Then
                  Cnt := Cnt + 1;

                if (((B16.Red > 220) and (B16.Blue > 220) and (B16.Green > 220)) Or
                    ((B16.Red + B16.Blue + B16.Green) > 600)) Then
                  Cnt := Cnt + 1;

                if Cnt >= 18 Then
                begin
                  A.Red   := 255;
                  A.Blue  := 255;
                  A.Green := 255;

                  Dst.Rgb[X, Y]     := A;
                end;

                if Cnt >= 23 Then
                begin
                  A.Red   := 255;
                  A.Blue  := 255;
                  A.Green := 255;

                  Dst.Rgb[X-1, Y-1] := A;
                  Dst.Rgb[X, Y-1]   := A;
                  Dst.Rgb[X+1, Y-1] := A;

                  Dst.Rgb[X-1, Y]   := A;
                  Dst.Rgb[X+1, Y]   := A;

                  Dst.Rgb[X-1, Y+1] := A;
                  Dst.Rgb[X, Y+1]   := A;
                  Dst.Rgb[X+1, Y+1] := A;
                end;

              end;
            end;
          end { if < 30}
          else
          begin
            if ((A.Red > 245) and (A.Blue > 245) and (A.Green > 245)) Or
                 ((A.Red + A.Blue + A.Green) > 735) then
            begin
              A.Red   := 255;
              A.Blue  := 255;
              A.Green := 255;
              Dst.Rgb[X, Y] := A;
            end
            else
            if UseAgresiveMode then
            begin
              if ((A.Red > 238) and (A.Blue > 238) and (A.Green > 238)) Or
                   ((A.Red + A.Blue + A.Green) > 715) then
              begin
                A.Red   := 255;
                A.Blue  := 255;
                A.Green := 255;
                Dst.Rgb[X, Y] := A;
              end
            end;
          end;
        end; { for Y}
      end; { for X}

      TmpDib.Assign(Dst);
    end;
  Finally
    FreeAndNil(TmpDib);
  End;

  GamaDib := TDibGraphic.Create;
  GamaTransform := TGammaTransform.Create;
  try
    GamaDib.Assign(Dst);
    GamaTransform.Gamma := 0.5;
    GamaTransform.ApplyOnDest(GamaDib, Dst);
  finally
    FreeAndNil(GamaDib);
    FreeAndNil(GamaTransform);
  end;
end;

function Tdm_Image.CleanDibBorder(SrcJpgDib: TJpegGraphic): Boolean;
Var
  CheckRect : TRect;
  CropRect  : TRect;
  TmpJPG    : TJpegGraphic;
  ChkDib    : TTiffGraphic;
  SrcDib    : TDibGraphic;
  DstDib    : TDibGraphic;
begin
  Result := False;
  Try
    TmpJPG := TJpegGraphic.Create;
    ChkDib := TTiffGraphic.Create;
    Try
      TmpJPG.Assign(SrcJpgDib);
      ChkDib.Assign(SrcJpgDib);

      SrcDib := TDibGraphic.Create;
      DstDib := TDibGraphic.Create;
      Try
        SrcDib.Assign(ChkDib);
        ConvertDibToBlackWhite(SrcDib,DstDib);
        ChkDib.Assign(DstDib);
      Finally
        FreeAndNil(SrcDib);
        FreeAndNil(DstDib);
      End;

      CheckRect.Left := 0;
      CheckRect.Right := SrcJpgDib.Width;
      CheckRect.Top := 0;
      CheckRect.Bottom := SrcJpgDib.Height;

      CropRect.Left := FindMostCleanLine(ChkDib, fcFullClean, fcLeftToRight, clFromBegining, CheckRect.Left);
      CropRect.Right := FindMostCleanLine(ChkDib, fcFullClean, fcRightToLeft,clFromBegining, CheckRect.Right);
      CropRect.Top := FindMostCleanLine(ChkDib, fcFullClean, fcTopToBottom, clFromBegining, CheckRect.Top);
      CropRect.Bottom := FindMostCleanLine(ChkDib, fcFullClean, fcBottomToTop, clFromBegining, CheckRect.Bottom);

      IF (CropRect.Left < CropRect.Right) And (CropRect.Top < CropRect.Bottom) Then
        DoCropDib(TmpJPG, SrcJpgDib, CropRect);
      Result := True;
    Finally
      FreeAndNil(ChkDib);
      FreeAndNil(TmpJPG);
    End;
  Except;
    Result := False;
  End;
end;

procedure Tdm_Image.ConvertWebBrowser2Jpg(const wb: TWebBrowser; const fileName: TFileName) ;
var
  viewObject : IViewObject;
  r : TRect;
  bitmap : TBitmap;
begin
  if wb.Document <> nil then
  begin
    wb.Document.QueryInterface(IViewObject, viewObject) ;
    if Assigned(viewObject) then
    try
      bitmap := TBitmap.Create;
      try
        r := Rect(0, 0, wb.Width, wb.Height) ;

        bitmap.Height := wb.Height;
        bitmap.Width := wb.Width;

        viewObject.Draw(DVASPECT_CONTENT, 1, nil, nil, Application.Handle, bitmap.Canvas.Handle, @r, nil, nil, 0) ;

        with TJPEGImage.Create do
        try
          Assign(bitmap) ;
          SaveToFile(fileName) ;
        finally
          Free;
        end;
      finally
        bitmap.Free;
      end;
    finally
      viewObject._Release;
    end;
  end;
end;

procedure Tdm_Image.ConvertDibToGray(SrcDib: TDibGraphic;
                                     DstDib: TDibGraphic);
var
  Transform: TConvertToGrayTransform;
begin
  Transform := TConvertToGrayTransform.Create;
  try
   Transform.ApplyOnDest(SrcDib, DstDib);
  finally
   FreeAndNil(Transform);
  end;
end;

procedure Tdm_Image.ConvertDibToGrayLzw(Dst: TDibGraphic);
Var
  TmpTif : TTiffGraphic;
  SrcRect: TRect;
  DstRect: TRect;
begin
  TmpTif := TTiffGraphic.Create;
  try
    TmpTif.NewImage(Dst.Width,Dst.Height,ifGray16, nil, Dst.XDotsPerInch, Dst.YDotsPerInch );
    TmpTif.Compression := tcLZW;

    SrcRect.Top    := 0;
    SrcRect.Left   := 0;
    SrcRect.Right  := Dst.Width;
    SrcRect.Bottom := Dst.Height;

    DstRect.Top    := 0;
    DstRect.Left   := 0;
    DstRect.Right  := Dst.Width;
    DstRect.Bottom := Dst.Height;

    TmpTif.Canvas.CopyRect(DstRect,
                            Dst.Canvas,
                            SrcRect);
    Dst.Assign(TmpTif);
  finally
    FreeAndNil(TmpTif);
  end;
end;

procedure Tdm_Image.ConvertTifToTifBW(SrcTifFile : WideString; DstTifFile : WideString);
Var
  SrcTif    : TTiffGraphic;
  CnvTif    : TTiffGraphic;
  TmpTif    : TTiffGraphic;
  CurFrame  : Integer;
  SrcRect   : TRect;
  DstRect   : TRect;

  procedure TransformTif(Src : TTiffGraphic; Dst : TTiffGraphic);
  Var
    SrcDib : TDibGraphic;
    DstDib : TDibGraphic;
  begin
    SrcDib := TDibGraphic.Create;
    DstDib := TDibGraphic.Create;
    Try
      SrcDib.Assign(Src);
      ConvertDibToBlackWhite(SrcDib,DstDib);
      Dst.Assign(DstDib);
    Finally
      FreeAndNil(SrcDib);
      FreeAndNil(DstDib);
    End;
  end;

  procedure AppendTifToFile(TmpTIF : TTiffGraphic);
  Var
    Stream   : TFileStream;
  begin
    Try
      IF Not FileExists(DstTifFile) Then
      begin
        Stream  := TFileStream.Create(DstTifFile, fmCreate);
        TmpTif.SaveToStream(Stream)
      end
      else
      begin
      Stream  := TFileStream.Create(DstTifFile, fmOpenReadWrite);
        TmpTif.AppendToStream(Stream);
      end;
    Finally
      FreeAndNil(Stream);
    End;
  end;
begin
  Windows.DeleteFile(PWideChar(DstTifFile));
  SrcTif := TTiffGraphic.Create;
  TmpTif := TTiffGraphic.Create;
  try
    SrcTif.MultiLoad := True;
    SrcTif.LoadFromFile(SrcTifFile);
    IF SrcTif.FrameCount = 0 Then
    begin
      TransformTif(SrcTif,TmpTif);
      AppendTifToFile(TmpTIF);
    end
    else
    begin
      CnvTif := TTiffGraphic.Create;
      Try
        For CurFrame := 1 To SrcTif.FrameCount Do
        begin
          CnvTif.Assign(SrcTif.Frames[CurFrame]);
          TransformTif(CnvTif,TmpTif);
          AppendTifToFile(TmpTIF);
        end;
      Finally
        FreeAndNil(CnvTif);
      End;
    end;
  Finally
    FreeAndNil(SrcTif);
    FreeAndNil(TmpTif);
  end;
end;

procedure Tdm_Image.ConvertPdfToTifBW(SrcPdfFile : WideString; DstTifFile : WideString);
Var
  Rslt         : Integer;
  PdfFileName  : WideString;
  PDFLibrary   : TQuickPDF;
  MemoryStream :TMemoryStream;
  TifFilename  : WideString;
  JpgFileNmae  : WideString;
  FHndl        : Integer;
  PageRef      : Integer;
  UTF8_String  : UTF8String;
  UnlockResult : Integer;
  TmpJpg       : TJPegGraphic;
  TmpTif       : TTiffGraphic;
  TmpDib       : TDibGraphic;
  CurCursor    : TCursor;
  CurPage      : Integer;
  PdfPageCount : Integer;

  procedure AppendTifToFile(TmpTIF : TTiffGraphic);
  Var
    Stream   : TFileStream;
  begin
    Try
      IF Not FileExists(DstTifFile) Then
      begin
        Stream  := TFileStream.Create(DstTifFile, fmCreate);
        TmpTif.SaveToStream(Stream)
      end
      else
      begin
      Stream  := TFileStream.Create(DstTifFile, fmOpenReadWrite);
        TmpTif.AppendToStream(Stream);
      end;
    Finally
      FreeAndNil(Stream);
    End;
  end;
begin
  Windows.DeleteFile(PWideChar(DstTifFile));
  PdfFileName := Trim(SrcPdfFile);
  if Not FileExists(PdfFileName) then
    Exit;

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
      try
        PdfPageCount := PDFLibrary.DAGetPageCount(FHndl);
        for CurPage := 1 to PdfPageCount do
        begin
          PageRef := PDFLibrary.DAFindPage(FHndl,CurPage);
          IF PageRef > 0 Then
          begin
            MemoryStream := TMemoryStream.Create;
            Try
              PDFLibrary.DARenderPageToStream(FHndl,
                                              PageRef,
                                              1 {Jpg},
                                              Dft_DPI_Render,
                                              MemoryStream {FileName});
              MemoryStream.Seek(0, soFromBeginning);
              TmpJpg := TJPegGraphic.Create;
              TmpTif := TTiffGraphic.Create;
              Try
                TmpJpg.LoadFromStream(MemoryStream);
                TmpDib := TDibGraphic.Create;
                Try
                  TmpDib.Assign(TmpJpg);
                  ConvertDibToGrayLzw(TmpDib);
                  TmpTif.Assign(TmpDib);
                Finally
                  FreeAndNil(TmpDib);
                  FreeAndNil(TmpTif);
                End;
              Finally
                FreeAndNil(TmpJpg);
              End;
            Finally
              FreeAndNil(MemoryStream);
            End;
          end;
        end;
      finally
        PDFLibrary.DACloseFile(FHndl);
      end;
    Finally
      Screen.Cursor := CurCursor;
    End;
  Finally
    FreeAndNil(PDFLibrary);
  End;
end;

procedure Tdm_Image.ConvertDibToBlackWhite(Src: TDibGraphic; Dst: TDibGraphic);
var
  fGraphic      : TTiffGraphic;
  SrcRect       : TRect;
  DstRect       : TRect;
  TmpDst        : TDibGraphic;
  Transform     : TImageFormatTransform;
  GamaTransform : TGammaTransform;
  GamaDib       : TDibGraphic;
begin
  Dst.Clear;
  Transform := TImageFormatTransform.Create;
  try
    Transform.ImageFormat := ifBlackWhite;
    Transform.Quantize    := True;
    Transform.Dither      := True;
    Transform.ApplyOnDest(Src, Dst);
  finally
    FreeAndNil(Transform);
  end;

  GamaDib := TDibGraphic.Create;
  GamaTransform := TGammaTransform.Create;
  try
    GamaDib.Assign(Dst);
    GamaTransform.Gamma := 0.5;
    GamaTransform.ApplyOnDest(GamaDib, Dst);
  finally
    FreeAndNil(GamaDib);
    FreeAndNil(GamaTransform);
  end;

  fGraphic := TTiffGraphic.Create;
  Try
   fGraphic.NewImage(Dst.Width, Dst.Height, ifBlackWhite, nil, 600, 600);
   fGraphic.Compression := tcGroup4;

   SrcRect.Top := 0;
   SrcRect.Left := 0;
   SrcRect.Right := Dst.Width;
   SrcRect.Bottom := Dst.Height;

   DstRect.Top := 0;
   DstRect.Left := 0;
   DstRect.Right := Dst.Width;
   DstRect.Bottom := Dst.Height;

   SetStretchBltMode(fGraphic.Canvas.Handle, STRETCH_DELETESCANS);
   fGraphic.Canvas.CopyRect(DstRect, Dst.Canvas, SrcRect);

   Dst.Assign(fGraphic);
  Finally
   FreeAndNil(fGraphic);
  End;
end;

procedure Tdm_Image.ConvertDibToBlackWhite(Dst: TDibGraphic);
Var
  Src: TDibGraphic;
begin
  Src := TTiffGraphic.Create;
  Try
   Src.Assign(Dst);
   ConvertDibToBlackWhite(Src,Dst);
  Finally
   FreeAndNil(Src);
  End;
end;

procedure Tdm_Image.ConvertDibGray2BW(Src: TDibGraphic; Dst: TDibGraphic);
Var
  TmpTif : TTiffGraphic;
  SrcRect: TRect;
  DstRect: TRect;
begin
  TmpTif := TTiffGraphic.Create;
  try
    TmpTif.NewImage(Src.Width,Src.Height,ifBlackWhite, nil, Src.XDotsPerInch, Src.YDotsPerInch );
    TmpTif.Compression := tcGroup4;

    SrcRect.Top    := 0;
    SrcRect.Left   := 0;
    SrcRect.Right  := Src.Width;
    SrcRect.Bottom := Src.Height;

    DstRect.Top    := 0;
    DstRect.Left   := 0;
    DstRect.Right  := Src.Width;
    DstRect.Bottom := Src.Height;

    TmpTif.Canvas.CopyRect(DstRect,
                            Src.Canvas,
                            SrcRect);
    Dst.Assign(TmpTif);
  finally
    FreeAndNil(TmpTif);
  end;
end;

procedure Tdm_Image.ConvertDib2BW(Src: TDibGraphic; Dst: TDibGraphic);
var
  X : Integer;
  Y : Integer;
  A : TRGB;
  TransformBW   : TImageFormatTransform;
  GamaTransform : TGammaTransform;
  DspcTransform : TDespeckleTransform;
  GamaDib       : TDibGraphic;
  BWDib         : TDibGraphic;
  aa : TBitMapGraphic;
  TmpSrc: TDibGraphic;
  TmpDst: TDibGraphic;
begin
  TmpSrc:= TDibGraphic.Create;
  TmpDst:= TDibGraphic.Create;
  Try
    TmpSrc.Assign(Src);

    ConvertDibToGrayLzw(TmpSrc);
    ConvertDibGray2BW(TmpSrc,TmpDst);

    TmpSrc.Assign(TmpDst);
    DspcTransform := TDespeckleTransform.Create;
    try
      DspcTransform.ApplyOnDest(TmpSrc, TmpDst);
      Src.Assign(Dst);
    finally
      FreeAndNil(DspcTransform);
    end;

    Dst.Assign(TmpDst);
  Finally
    FreeAndNil(TmpSrc);
    FreeAndNil(TmpDst);
  End;
end;

procedure Tdm_Image.ConvertDib2BitMap(Src: TDibGraphic; Dst: TBitMap);
Var
  TmpBarCode: TBitmapGraphic;
begin
  TmpBarCode := TBitmapGraphic.Create;
  Try
    TmpBarCode.Assign(Src);
    Dst.Assign(TmpBarCode);
  Finally
    FreeAndNil(TmpBarCode);
  End;
end;

function Tdm_Image.ImgToBase64(FileName : String) : String;
Var
  fMM : TMemoryStream;
  fSS : TStringStream;
  base64: TBase64Encoding;
begin
  Try
    fMM := TMemoryStream.Create;
    fSS := TStringStream.Create('');

    fMM.LoadFromFile(FileName);
    base64 := TBase64Encoding.Create(0); //0: Do not add carriage return and line feed character
    base64.Encode(fMM, fSS);
    Result := fSS.DataString;
  Finally
    FreeAndNil(fMM);
    FreeAndNil(fSS);
  End;
end;

function  Tdm_Image.Base64ToImg(Base64Str : String; DibGraphic : TJpegGraphic) : Boolean;
Var
  fMM : TMemoryStream;
  fSS : TStringStream;
  base64: TBase64Encoding;
  TempJpgFileName : String;
begin
  Result := False;
  DibGraphic.Clear;

  Try
    fMM := TMemoryStream.Create;
    fSS := TStringStream.Create(Base64Str);

    base64 := TBase64Encoding.Create(0); //0: Do not add carriage return and line feed character
    base64.Decode(fSS, fMM);

    TempJpgFileName := RemoveBackSlashChar(ParamArea.WinTempDir) + '\' +
                                  CleanDirFileName(PrmUserName) + '_1' +
                                 IntToStr(GetTickCount) + '.Jpg';

    fMM.SaveToFile(TempJpgFileName);
    DibGraphic.LoadFromFile(TempJpgFileName);
    Result := not DibGraphic.IsEmpty;
  Finally
    FreeAndNil(fMM);
    FreeAndNil(fSS);
  End;

end;

function Tdm_Image.FindMostCleanLine(SrcDib: TDibGraphic;
                                     FindCleanLineMode: TFindCleanLineMode;
                                     FindCleanLineDirection: TFindCleanLineDirection;
                                     CleanLineSearchMethode: TCleanLineSearchMethode;
                                     FromPos: Integer): Integer;
Var
  X, Y: Integer;
  A: TRGB;
  ToY: Integer;
  ToX: Integer;
  BX: Integer;
  LastBX: Integer;
  RowsFound  : Integer;
  PrecentCheck : Double;
Const
  CleanLessThen  : Integer = 20;
  CleanBlankLine : Integer = 5;
  CleanRowsLine  : Integer = 5;
begin

  PrecentCheck := FindMostCleanPrecect / 100;
  IF FindCleanLineMode = fcDerty Then
    PrecentCheck := FindMostDertyCleanPrecect / 100;

  IF FromPos = 0 Then
  begin
    Case FindCleanLineDirection Of
      fcLeftToRight: FromPos := 0;
      fcRightToLeft: FromPos := SrcDib.Width - 1;
      fcTopToBottom: FromPos := 0;
      fcBottomToTop: FromPos := SrcDib.Height - 1;
    End;
  end
  else
  begin
    Case FindCleanLineDirection Of
     fcLeftToRight: IF FromPos < 0 Then
                      FromPos := 0;
     fcRightToLeft: IF FromPos >= SrcDib.Width Then
                      FromPos := SrcDib.Width - 1;
     fcTopToBottom: IF FromPos < 0 Then
                      FromPos := 0;
     fcBottomToTop: IF FromPos >= SrcDib.Height Then
                      FromPos := SrcDib.Height - 1;
    End;
  end;
  Result := FromPos;

  RowsFound := 0;
  IF FindCleanLineDirection = fcLeftToRight Then
  Begin
    LastBX := 999999999;
    ToX := SrcDib.Width - 1;

    for X := FromPos To ToX do
    begin
      BX := 0;
      for Y := 0 to SrcDib.Height - 1 do
      begin
        Try
          A := SrcDib.Rgb[X, Y];
          if ((A.Red > FindCleanLine_Min_Red) Or
              (A.Blue > FindCleanLine_Min_Blue) Or
              (A.Green > FindCleanLine_Min_Green)) And
             ((A.Red + A.Blue + A.Green) > FindMostCleanWhiteTotalSum) then
          begin
            BX := BX + 1;
          end; { if }
        Except;
        End;
      end; { for }

      IF BX > PrecentCheck * SrcDib.Height Then
      begin
        RowsFound := RowsFound + 1;
      end
      else
      begin
        RowsFound := 0;
      end;

      IF RowsFound > CleanRowsLine Then
      begin
        Result := X - CleanRowsLine;
        if Result < 0 then
          Result := 0;
        Break;
      end;
    end; { for }
  End
  else IF FindCleanLineDirection = fcRightToLeft Then
  Begin
    LastBX := 999999999;
    ToX := 0;

    for X := FromPos DownTo ToX do
    begin
      BX := 0;
      for Y := 0 to SrcDib.Height - 1 do
      begin
        Try
          A := SrcDib.Rgb[X, Y];
          if ((A.Red > FindCleanLine_Min_Red) Or
              (A.Blue > FindCleanLine_Min_Blue) Or
              (A.Green > FindCleanLine_Min_Green)) And
              ((A.Red + A.Blue + A.Green) > FindMostCleanWhiteTotalSum) then
          begin
            BX := BX + 1;
          end; { if }
        Except;
        End;
      end; { for }

      IF BX > PrecentCheck * SrcDib.Height Then
      begin
        RowsFound := RowsFound + 1;
      end
      else
      begin
        RowsFound := 0;
      end;

      IF RowsFound > CleanRowsLine Then
      begin
        Result := X + CleanRowsLine;
        if Result > SrcDib.Width then
          Result := SrcDib.Width;
        Break;
      end;
    end; { for }
  End
  else IF FindCleanLineDirection = fcTopToBottom Then
  Begin
    LastBX := 999999999;
    ToY := SrcDib.Height - 1;

    for Y := FromPos To ToY do
    begin
      BX := 0;
      for X := 0 to SrcDib.Width - 1 do
      begin
        Try
          A := SrcDib.Rgb[X, Y];
          if ((A.Red > FindCleanLine_Min_Red) Or
              (A.Blue > FindCleanLine_Min_Blue) Or
              (A.Green > FindCleanLine_Min_Green)) And
              ((A.Red + A.Blue + A.Green) > FindMostCleanWhiteTotalSum) then
          begin
            BX := BX + 1;
          end; { if }
        Except;
        End;
      end; { for }

      IF BX > PrecentCheck * SrcDib.Width Then
      begin
        RowsFound := RowsFound + 1;
      end
      else
      begin
        RowsFound := 0;
      end;

      IF RowsFound > CleanRowsLine Then
      begin
        Result := Y - CleanRowsLine;
        if Result < 0 then
          Result := 0;
        Break;
      end;
    end; { for }
  End
  else IF FindCleanLineDirection = fcBottomToTop Then
  Begin
    LastBX := 999999999;
    ToY := 0;
    for Y := FromPos Downto ToY do
    begin
      BX := 0;
      for X := 0 to SrcDib.Width - 1 do
      begin
        Try
          A := SrcDib.Rgb[X, Y];
          if ((A.Red > FindCleanLine_Min_Red) Or
              (A.Blue > FindCleanLine_Min_Blue) Or
              (A.Green > FindCleanLine_Min_Green)) And
              ((A.Red + A.Blue + A.Green) > FindMostCleanWhiteTotalSum) then
          begin
           BX := BX + 1;
          end; { if }
        Except;
        End;
      end; { for }

      IF BX > PrecentCheck * SrcDib.Width Then
      begin
        RowsFound := RowsFound + 1;
      end
      else
      begin
        RowsFound := 0;
      end;

      IF RowsFound > CleanRowsLine Then
      begin
        Result := Y + CleanRowsLine;
        if Result > SrcDib.Height then
          Result := SrcDib.Height;
        Break;
      end;

    end; { for }
  End;
end;

function Tdm_Image.DoCropDib(SrcDib: TDibGraphic;
                              DstDib: TDibGraphic;
                              CropRect: TRect): Boolean;
var
  Transform: TCropTransform;
begin
  Result := False;
  Try
    Transform := TCropTransform.Create;
    try
     Transform.CropMode := cmExtractRect;
     Transform.Left     := CropRect.Left;
     Transform.Right    := CropRect.Right;
     Transform.Top      := CropRect.Top;
     Transform.Bottom   := CropRect.Bottom;

     DstDib.Assign(SrcDib);
     IF (Transform.Left <= Transform.Right) And
       (Transform.Top <= Transform.Bottom) Then
      Transform.ApplyOnDest(SrcDib, DstDib);
     Result := true;
    finally
     FreeAndNil(Transform);
    end;
  Except;
   Result := False;
  End;
end;

function Tdm_Image.RenderFileToJpgList(Pdf_FileName: String;
                                       ToPath : String;
                                       FileList: TStringList): Boolean;
Var
  Rslt         : Integer;
  UnlockResult : Integer;
  PDFLibrary   : TQuickPDF;
  t1           : Integer;
  FHndl        : Integer;
  PageRef      : Integer;
  CurPage      : Integer;
  TotalPages   : Integer;
  JpegImage    : TJpegImage;
  //Bitmap       : TBitmap;
  JpgFileName  : String;
  CurCursor    : TCursor;
begin
  Result := False;
  FileList.Clear;
  t1 := GetTickCount;
  TotalPages := -1;

  PDFLibrary := TQuickPDF.Create;
  Try
    UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
    if UnlockResult <> 1 then
      Exit;

    CurCursor := Screen.Cursor;
    Try
      Screen.Cursor := crHourGlass;
      FHndl := PDFLibrary.DAOpenFileReadOnly(Pdf_FileName,'');
      IF FHndl > 0 Then
      begin
        TotalPages := PDFLibrary.DAGetPageCount(FHndl);
        For CurPage := 1 To TotalPages Do
        begin
          PageRef := PDFLibrary.DAFindPage(FHndl, CurPage);
          IF PageRef > 0 Then
          begin
            JpgFileName := RemoveBackSlashChar(ToPath) + '\TmpJpg' + IntToStr(t1) + '_' + ZeroIntToStr(4,CurPage) + '.jpg';
            Rslt := PDFLibrary.DARenderPageToFile(FHndl,
                                                  PageRef,
                                                  1{Jpg},
                                                  ParamArea.Param_ConvertToPDF_DPI,
                                                  JpgFileName);
            FileList.Add(JpgFileName);
          end;
        end;
      End;
    Finally
      Screen.Cursor := CurCursor;
    End;
  Finally
    PDFLibrary.DACloseFile(FHndl);
    FreeAndNil(PDFLibrary);
  End;

  Result := (FileList.Count = TotalPages);
end;

function Tdm_Image.RenderPDFToMultiTiff(Pdf_FileName: String;
                                        Tmp: TTiffGraphic): Boolean;
Var
  Rslt         : Integer;
  PdfFileName  : WideString;
  PDFLibrary   : TQuickPDF;
  MemoryStream : TMemoryStream;
  TifFilename  : WideString;
  JpgFileNmae  : WideString;
  FHndl        : Integer;
  PageRef      : Integer;
  UTF8_String  : UTF8String;
  UnlockResult : Integer;
  TmpJpg       : TJPegGraphic;
  SaveTif      : TTiffGraphic;
  CurCursor    : TCursor;
  PdfPageCount : Integer;
  PageRane     : String;
  DstFileName  : String;
  DstPath      : String;
  NewFileName  : String;
  FileList     : TStringList;
  CurJpg       : Integer;
  Stream       : TMemoryStream;
begin
  Result := False;
  Tmp.Clear;
  PdfFileName := Trim(Pdf_FileName);
  if Not FileExists(PdfFileName) then
    Exit;

  DstPath := GetPdfMergeTmpPath;
  NewFileName := DstPath + '\TmpTif.Tif';
  Windows.DeleteFile(PWideChar(NewFileName));

  PdfPageCount := -1;
  FileList := TStringList.Create;
  Try
    If RenderFileToJpgList(PdfFileName,DstPath,FileList) Then
    begin
      PdfPageCount := FileList.Count;
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
        FreeAndNil(TmpJpg);
      End;
      Stream.Seek(0, soFromBeginning);
      Tmp.MultiLoad := True;
      Tmp.LoadFromStream(Stream);
    end;
    Result := Not Tmp.IsEmpty;
    if Tmp.FrameCount = 0 then
      Result := (PdfPageCount = 1)
    else
      Result := (Tmp.FrameCount = PdfPageCount);
  Finally
    if Assigned(SaveTif) then
      FreeAndNil(SaveTif);
    if Assigned(Stream) then
      FreeAndNil(Stream);
    FreeAndNil(FileList);
  End;
end;

function Tdm_Image.RenderFirstPageToTif(Pdf_FileName: String;
                                         Tif_FileName: String): Boolean;
Var
  PDFLibrary: TQuickPDF;
  UnlockResult: Integer;
  FHndl, Rslt, Page, PageRef, Options: Integer;
  CurCursor: TCursor;
  DestPath : String;
begin
  Page := 1;
  Options := 7;
  Result := False;

  DestPath := j_PathFileName(Tif_FileName);
  CurCursor := Screen.Cursor;
  Try
    Screen.Cursor := crHourGlass;
    DestPath := j_PathFileName(Tif_FileName);
    ForceDirectories(DestPath);

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
                                                  Dft_DPI_Render, // not ParamArea.Param_ConvertToPDF_DPI,
                                                                  // it just to view first page
                                                  Tif_FileName);
           Result := (Rslt = 1);
          end;
        Finally
          PDFLibrary.DACloseFile(FHndl);
        End;
      end
      else
      begin
        Rslt := PDFLibrary.LastErrorCode;
        Dialogs.ShowMessage(IntToStr(Rslt));
      end;
    Finally
      FreeAndNil(PDFLibrary);
    End;
  Finally
   Screen.Cursor := CurCursor;
  End;
end;

function Tdm_Image.RenderPdfPageToStream(FileName : WideString; PageNo : Integer; DPI : Integer; MS : TMemoryStream) : Boolean;
Var
  PDFLibrary   : TQuickPDF;
  FHndl        : Integer;
  PageRef      : Integer;
  UnlockResult : Integer;
  Rslt         : Integer;
  PathName     : WideString;
  TmpFileName  : WideString;
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

function Tdm_Image.RenderFirstPageToDib(Pdf_FileName: String;
                                        Tmp: TDibGraphic): Boolean;
Var
  Rslt         : Integer;
  PdfFileName  : WideString;
  PDFLibrary   : TQuickPDF;
  MemoryStream :TMemoryStream;
  TifFilename  : WideString;
  JpgFileNmae  : WideString;
  FHndl        : Integer;
  PageRef      : Integer;
  UTF8_String  : UTF8String;
  UnlockResult : Integer;
  TmpJpg       : TJPegGraphic;
  CurCursor    : TCursor;
begin
  Result := False;
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
          PageRef := PDFLibrary.DAFindPage(FHndl,1);
          IF PageRef > 0 Then
          begin
            MemoryStream := TMemoryStream.Create;
            Try
              PDFLibrary.DARenderPageToStream(FHndl,
                                              PageRef,
                                              1 {Jpg},
                                              Dft_DPI_Render,
                                              MemoryStream {FileName});
              MemoryStream.Seek(0, soFromBeginning);
              TmpJpg := TJPegGraphic.Create;
              Try
                TmpJpg.LoadFromStream(MemoryStream);
                Tmp.Assign(Tmpjpg);
                Result := True; // success
              Finally
                FreeAndNil(TmpJpg);
              End;
            Finally
              FreeAndNil(MemoryStream);
            End;
          end;
        end;
      Finally
        Screen.Cursor := CurCursor;
      End;
    Finally
      FreeAndNil(PDFLibrary);
    End;
  Except;
    Result := False; // error on proc
  End;
end;

function Tdm_Image.ConvertWordToPDF(FileName : String; ToFileName : String) : Boolean;
begin
  Result:= Word_ConvertFileToPDF(FileName, ToFileName, False, False);
end;

function Tdm_Image.ConvertWordToRTF(FileName : String; ToFileName : String) : Boolean;
begin
  Result := ConvertWord2RTF(FileName, ToFileName);
end;

function Tdm_Image.ConvertPdfToTif(Pdf_FileName: String;
                                    DstFileName: String;
                                    ConvertTo_BW: Boolean = true;
                                    ConvertJustFirstPage: Boolean = False): Boolean;
Var
  PDFLibrary: TQuickPDF;
  FHndl: Integer;
  PageRef: Integer;
  UnlockResult: Integer;
  fTempDir: String;
  TmpTif: TTiffGraphic;
  TmpJpeg: TJpegGraphic;
  CurFrame: Integer;
  ToFileName: String;
  DestPath  : String;
  Rslt: Integer;
  CurCursor: TCursor;
  PdfPageCount : Integer;
  TiffColorType : TTiffColorType;
begin
  Result := False;
  fTempDir := ParamArea.WinTempDir;
  TiffColorType := ctColor;
  If ConvertTo_BW Then
    TiffColorType := ctBW;
  IF FileExists(DstFileName) Then
   DeleteFile(DstFileName);
  ToFileName := ChangeFileExt(DstFileName, '.Jpg');
  ToFileName := RemoveBackSlashChar(fTempDir) + '\' + j_FileName(ToFileName);

  PDFLibrary := TQuickPDF.Create;
  Try
    UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
    IF UnlockResult <> 1 then
      Exit;

    IF Not FileExists(Pdf_FileName) Then
      Exit;

    CurCursor := Screen.Cursor;
    Try
      Screen.Cursor := crHourGlass;
      FHndl := PDFLibrary.DAOpenFileReadOnly(Pdf_FileName, '');
      IF FHndl > 0 Then
      begin
        Try
          PdfPageCount := PDFLibrary.DAGetPageCount(FHndl);
          For CurFrame := 1 to PdfPageCount Do
          begin
            ToFileName := RemoveBackSlashChar(fTempDir) + '\Tmp^Cnv.Jpg';
            DestPath := j_PathFileName(ToFileName);
            ForceDirectories(DestPath);

            PageRef := PDFLibrary.DAFindPage(FHndl, CurFrame);
            IF PageRef > 0 Then
            begin
              PDFLibrary.AddTrueTypeFont(dm_SiLang.siLangLinked.GetTextOrDefault('IDS_31' (* 'DEFAULT_CHARSET' *) ), 0);
              Rslt := PDFLibrary.DARenderPageToFile(FHndl, PageRef, 1,
              ParamArea.Param_ConvertToPDF_DPI, ToFileName);
              IF (Rslt = 1) Then
              begin
                TmpJpeg := TJpegGraphic.Create;
                TmpTif := TTiffGraphic.Create;
                Try
                  TmpJpeg.LoadFromFile(ToFileName);
                  TmpTif.Assign(TmpJpeg);
                  AppendTifToFile(TmpTif, DstFileName, ParamArea.Param_ConvertToPDF_DPI,
                    ParamArea.Param_ConvertToPDF_DPI, TiffColorType);
                Finally
                  FreeAndNil(TmpJpeg);
                  FreeAndNil(TmpTif);
                End;
              end;
            end
            else
            begin
             IF FileExists(DstFileName) Then
              DeleteFile(DstFileName);
             Break;
            end;
            if ConvertJustFirstPage then
              Break;
          end;

          if Not ConvertJustFirstPage then
          begin
            Result := FileExists(DstFileName) And
                     (PdfPageCount = GetImagePageCount(DstFileName));
          end
          else
          begin
            Result := FileExists(DstFileName) And
                     (GetImagePageCount(DstFileName) = 1);
          end;
        Finally
          PDFLibrary.DACloseFile(FHndl);
        End;
       end
       else
       begin
        Result := False;
       end;
    Finally
      Screen.Cursor := CurCursor;
    End;
  Finally
    FreeAndNil(PDFLibrary);
  End;
end;

function Tdm_Image.ConvertPdfToMultiTif(Pdf_FileName: String;
                                        DstFileName: String;
                                        ConvertTo_BW: Boolean = true): Boolean;
Var
  PDFLibrary: TQuickPDF;
  FHndl: Integer;
  PageRef: Integer;
  UnlockResult: Integer;
  fTempDir: String;
  PageRane: String;
  ColorImage: Integer;
  OutputOptions: Integer;
  TmpTif: TTiffGraphic;
  TmpJpeg: TJpegGraphic;
  CurFrame: Integer;
  ToFileName: String;
  DestPath  : String;
  DllFileName: String;
  Rslt: Integer;
  CurCursor: TCursor;
  PdfPageCount : Integer;
  TiffColorType : TTiffColorType;
begin
  Result := False;
  fTempDir := ParamArea.WinTempDir;
  ColorImage := 1;
  If Not ConvertTo_BW Then
    ColorImage := 0;
  OutputOptions:= 0;
  IF FileExists(DstFileName) Then
   DeleteFile(DstFileName);

  DestPath := j_PathFileName(DstFileName);
  ForceDirectories(DestPath);

  PDFLibrary := TQuickPDF.Create;
  Try
    UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
    IF UnlockResult <> 1 then
     Exit;

    IF Not FileExists(Pdf_FileName) Then
     Exit;

    CurCursor := Screen.Cursor;
    Try
      Screen.Cursor := crHourGlass;
      Rslt := PDFLibrary.LoadFromFile(Pdf_FileName, '');
      IF Rslt = 1 Then
      begin
        Try
          PdfPageCount := PDFLibrary.PageCount;
          PageRane := '1-' + IntToStr(PdfPageCount);
          Rslt := PDFLibrary.RenderAsMultipageTIFFToFile(ParamArea.Param_ConvertToPDF_DPI,
                                                         PageRane,
                                                         ColorImage,
                                                         OutputOptions,
                                                         DstFileName);

          if (Rslt = 1) then
          begin
            Result := FileExists(DstFileName);
          end;
        Finally;
        End;
      end
      else
      begin
        Result := False;
      end;
    Finally
      Screen.Cursor := CurCursor;
    End;
  Finally
    FreeAndNil(PDFLibrary);
  End;
end;

function Tdm_Image.ConvertTifToPdf(Tif_FileName: String;
                                    DstFileName: String): Boolean;
Var
  Rslt : Integer;
begin
  Rslt := AddMultiImgToPdfFile(Tif_FileName, DstFileName, False {FitToPage}, -1);
  Result := (Rslt = QuickPdf_NoError);
end;

function Tdm_Image.ConvertOfficeTo_(Office_FileName: String;
                                     Wanted_FileName: String;
                                     ToType: Integer;
                                     Var ErrMsg: String): Boolean;
var
  WFalse, WFormat, WFileName: OleVariant;
  Wrd: TWordApplication;
  wrddoc: Word2000.TWordDocument;
  CurCursor: TCursor;
  t1: Integer;

  function OpenDoc(Word: TWordApplication; FName: WideString): _Document;
  var
   FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, Revert, Format,
     Visible: OleVariant;
  begin
   FileName := FName;
   ConfirmConversions := False;
   ReadOnly := true;
   AddToRecentFiles := False;
   Revert := true;
   Format := wdOpenFormatAuto;
   Visible := False;
   Result := Word.Documents.Open(FileName, ConfirmConversions, ReadOnly,
    AddToRecentFiles, EmptyParam, EmptyParam, Revert, EmptyParam, EmptyParam,
    Format, EmptyParam, Visible{, EmptyParam, EmptyParam, EmptyParam});
  end;

begin
  Result := False;
  ErrMsg := '';
  WFalse := False;
  CurCursor := Screen.Cursor;
  t1 := GetTickCount;
  Try
    Screen.Cursor := crHourGlass;
    Try
      Wrd:=TWordApplication.Create(nil);
      try
        Wrd.ConnectKind := ckNewInstance;
        Wrd.AutoQuit := true;
        Wrd.Connect;
        wrddoc:= Word2000.TWordDocument.Create(nil);
        try
          wrddoc.ConnectKind := ckAttachToInterface;
          wrddoc.ConnectTo(OpenDoc(Wrd, Office_FileName));
          Sleep(10);
          WFormat := ToType;
          WFileName := Wanted_FileName;
          wrddoc.SaveAs(WFileName, WFormat);
          Sleep(10);
          wrddoc.Close;
          Wrd.Disconnect;
          Wrd.Quit(WFalse);
          Result := true;
        finally
          FreeAndNil(wrddoc);
        end;
      finally
        FreeAndNil(Wrd);
      end;
    Except
      On e: Exception Do
      begin
        Result := False;

        dm_Main.MzWriteLog(e.Message,MzLog_FromProcess);
        ErrMsg := e.Message;
      end;
    End;
  Finally
   Screen.Cursor := CurCursor;
  End;
end;

function Tdm_Image.ConvertPPT_ToRTF(PPT_FileName: String;
                                    Wanted_FileName: String;
                                    Var ErrMsg: String): Boolean;
var
  PowerPoint : TPowerPointApplication;
  VisFalse: OleVariant;
  WindowMin: OleVariant;
  EmbedFonts: OleVariant;
begin
  Result := False;
  Try
    Try
      LockWindowUpdate(Application.MainForm.Handle);
      PowerPoint := TPowerPointApplication.Create(Application);
      VisFalse := True;
      PowerPoint.Visible := VisFalse;
      WindowMin:= 2 {wsMinimized};
      PowerPoint.WindowState := WindowMin;
      PowerPoint.Presentations.Open(PPT_FileName, msoFalse, msoFalse, msoTrue);
      WindowMin:= 2 {wsMinimized};
      PowerPoint.WindowState := WindowMin;
      EmbedFonts := False;
      PowerPoint.ActivePresentation.SaveAs(Wanted_FileName, ppSaveAsRTF, EmbedFonts);
      PowerPoint.ActivePresentation.Close;
      Result := True;
    Finally
      PowerPoint.Quit;
      PowerPoint := nil;
      LockWindowUpdate(0);
    End;
  Except;
    Result := False;
  End;
end;

function Tdm_Image.ConvertRTFToHTML(RTF_FileName: String;
                                     Html_FileName: String;
                                     Var ErrMsg: String): Boolean;
begin
  Result := ConvertOfficeTo_(RTF_FileName, Html_FileName, mzFormatHTML, ErrMsg);
end;

function  Tdm_Image.ConvertOfficeToHtml(OfficeFileName : String; HTMLFileName : String) : Boolean;
Var
  Cnvr    : IHtDocumentConverter;
  HtmlStr : string;
begin
  Result := False;
  DeleteFile(HTMLFileName);
  if IsFileHTML(OfficeFileName) then
  begin
    Result := CopyFile(System.PWideChar(OfficeFileName),System.PWideChar(HTMLFileName),False);
    Exit;
  end;

  Try
    Cnvr := THtDocumentConverter.GetConverter(J_ExtFileName(OfficeFileName));

    if Cnvr = nil then
    begin
      Exit;
    end;

    try
      HtmlStr := Cnvr.ConvertFile(OfficeFileName);
      HtmlStr := '<style>body {background: white; margin: 16px}</style>' + HtmlStr;
      HtmlStr := ' <meta charset="UTF-8"> ' + HtmlStr;
      HtmlStr := ' <meta http-equiv="content-language" content="en" >  ' + HtmlStr;

      TFile.WriteAllText(HTMLFileName, UTF8Encode(HtmlStr));
    Except;
      HtmlStr := '<HTML> <body> <div>' +
                 'Fail convert - ' + OfficeFileName + ' To HTML ' +
                 '</div></body></HTML>';
      HtmlStr := '<style>body {background: white; margin: 16px}</style>' + HtmlStr;
      HtmlStr := ' <meta charset="UTF-8"> ' + HtmlStr;
      HtmlStr := ' <meta http-equiv="content-language" content="en" >  ' + HtmlStr;
    end;
  Finally
    Cnvr := nil;
  End;

  Result := FileExists(HTMLFileName);
end;

function  Tdm_Image.ConvertOutlookMsgHtml(MsgFileName : String; HTMLFileName : String) : Boolean;
var
  HtmlStr : String;
  Cnvr    : IHtDocumentConverter;
  CurChar : Integer;
begin
  Result := False;
  DeleteFile(HTMLFileName);
  Try
    Cnvr := THtDocumentConverter.GetConverter('msg');
    if Cnvr = nil then
    begin
      exit
    end;

    HtmlStr := Cnvr.ConvertFile(MsgFileName);
    HtmlStr := '<style>body {background: white; margin: 16px}</style>' + HtmlStr;
    HtmlStr := ' <meta charset="UTF-8"> ' + HtmlStr;
    HtmlStr := ' <meta http-equiv="content-language" content="en" >  ' + HtmlStr;

    for CurChar := 1 To Length(HtmlStr) Do
    begin
      case ord(HtmlStr[CurChar]) of
        224 : HtmlStr[CurChar] := 'א';
        225 : HtmlStr[CurChar] := 'ב';
        226 : HtmlStr[CurChar] := 'ג';
        227 : HtmlStr[CurChar] := 'ד';
        228 : HtmlStr[CurChar] := 'ה';
        229 : HtmlStr[CurChar] := 'ו';
        230 : HtmlStr[CurChar] := 'ז';
        231 : HtmlStr[CurChar] := 'ח';
        232 : HtmlStr[CurChar] := 'ט';
        233 : HtmlStr[CurChar] := 'י';
        234 : HtmlStr[CurChar] := 'ך';
        235 : HtmlStr[CurChar] := 'כ';
        236 : HtmlStr[CurChar] := 'ל';
        237 : HtmlStr[CurChar] := 'ם';
        238 : HtmlStr[CurChar] := 'מ';
        239 : HtmlStr[CurChar] := 'ן';
        240 : HtmlStr[CurChar] := 'נ';
        241 : HtmlStr[CurChar] := 'ס';
        242 : HtmlStr[CurChar] := 'ע';
        243 : HtmlStr[CurChar] := 'ף';
        244 : HtmlStr[CurChar] := 'פ';
        245 : HtmlStr[CurChar] := 'ץ';
        246 : HtmlStr[CurChar] := 'צ';
        247 : HtmlStr[CurChar] := 'ק';
        248 : HtmlStr[CurChar] := 'ר';
        249 : HtmlStr[CurChar] := 'ש';
        250 : HtmlStr[CurChar] := 'ת';
      end;
    end;

    TFile.WriteAllText(HTMLFileName, UTF8Encode(HtmlStr));
  Finally
    Cnvr := nil;
  End;

  Result := FileExists(HTMLFileName);
end;

function  Tdm_Image.ConvertOutlookMsg2Html(MsgFileName : String; out HTMLStr : String) : Boolean;
Const
  olHTML = $00000005;
Var
  Outlook : OLEVariant;
  TmpFileName : String;
begin
  Result := False;
  Try
    Try
      outlook := unassigned;
      try
        Outlook := GetActiveOleObject('Outlook.Application') ;
      except
        Try
          outlook := CreateOleObject('Outlook.Application');
        except
          on e:exception do
          begin
            JustWriteToLog('Fail to create Outlook OLE Object : Error, ' + e.Message);
            JustWriteToLog(' Can not SaveAs(TmpFileName) As olHTML : File - ' + MsgFileName);
            outlook := unassigned;
          end;
        end;
      end;

      if not VarIsEmpty(outlook) Then
      begin
        TmpFileName := RemoveBackSlashChar(ParamArea.WinTempDir) + '\Tmp' + IntToStr(GetTickCount) + '.html';
        Outlook := outlook.CreateItemFromTemplate(MsgFileName);
        Outlook.SaveAs(TmpFileName, olHTML);
        Result := True;
      end;
    Finally
      outlook := unassigned;
    End;

    if FileExists(TmpFileName) then
    begin
      TFile.ReadAllText(TmpFileName);
    end;
  Except
    On e: Exception Do
    begin
      JustWriteToLog('Error On ConvertOutlookMsg2Html : ' + e.Message);
    end;
  End;
end;

procedure Tdm_Image.EncryptBMP(const FileName: String;
                                Key: Integer);
Var
  TmpJPG: TJpegGraphic;
  TmpTif: TTiffGraphic;
  EnccryptTif: TTiffGraphic;
  TmpBmp1: TBitmapGraphic;
  TmpBMP: TBitmap;
  CurFrame: Integer;
begin
  TmpTif := TTiffGraphic.Create;
  TmpBMP := TBitmap.Create;
  TmpBmp1 := TBitmapGraphic.Create;
  Try
    LoadImageToTif(FileName, TmpTif);
    IF IsFileBMP(FileName) Then
    begin
      TmpBmp1.Assign(TmpTif);
      TmpBMP.Assign(TmpBmp1);
      EncryptBMP(TmpBMP, Key);
      TmpBMP.SaveToFile(FileName);
    end
    else IF IsFileTiff(FileName) Then
    begin
      IF TmpTif.FrameCount = 0 Then
      begin
        TmpBmp1.Assign(TmpTif);
        TmpBMP.Assign(TmpBmp1);
        EncryptBMP(TmpBMP, Key);
        TmpBmp1.Assign(TmpBMP);
        TmpTif.Assign(TmpBmp1);
        TmpTif.SaveToFile(FileName);
      end
      else
      begin
        EnccryptTif := TTiffGraphic.Create;
        Try
          For CurFrame := 1 To TmpTif.FrameCount Do
          begin
            TmpBmp1.Assign(TmpTif.Frames[CurFrame]);
            TmpBMP.Assign(TmpBmp1);
            EncryptBMP(TmpBMP, Key);
            TmpBmp1.Assign(TmpBMP);
            EnccryptTif.Assign(TmpBmp1);
            IF CurFrame = 1 Then
             SaveTifToFile(EnccryptTif, FileName)
            else
             AppendTifToFile(EnccryptTif, FileName,ctColor);
          end;
        Finally
          FreeAndNil(EnccryptTif);
        End;
      end;
    end
    else IF IsFileJpeg(FileName) Then
    begin
      TmpBmp1.Assign(TmpTif);
      TmpBMP.Assign(TmpBmp1);
      EncryptBMP(TmpBMP, Key);
      TmpJPG := TJpegGraphic.Create;
      Try
        TmpBmp1.Assign(TmpBMP);
        TmpJPG.Assign(TmpBmp1);
        TmpJPG.SaveToFile(FileName);
      Finally
        FreeAndNil(TmpJPG);
      End;
    end
    else
    begin
      TmpBmp1.Assign(TmpTif);
      TmpBMP.Assign(TmpBmp1);
      EncryptBMP(TmpBMP, Key);
      TmpBmp1.Assign(TmpBMP);
      TmpTif.Assign(TmpBmp1);
      TmpTif.SaveToFile(FileName);
    end;
  Finally
   FreeAndNil(TmpTif);
   FreeAndNil(TmpBMP);
   FreeAndNil(TmpBmp1);
  End;
end;

procedure Tdm_Image.EncryptBMP(const BMP: TBitmap;
                                Key: Integer);
var
  BytesPorScan: Integer;
  w, H: Integer;
  p: pByteArray;
begin
  try
   BytesPorScan := Abs(Integer(BMP.ScanLine[1]) - Integer(BMP.ScanLine[0]));
  except
   raise EAppError.Create('Error');
  end;
  RandSeed := Key;
  for H := 0 to BMP.Height - 1 do
  begin
   p := BMP.ScanLine[H];
   for w := 0 to BytesPorScan - 1 do
    p^[w] := p^[w] xor Random(256);
  end;
end;

procedure Tdm_Image.EncryptDIB(const FileName: String;
                                Key: Integer);
Var
  TmpBMP: TBitmapGraphic;
  TmpJPG: TJpegGraphic;
  TmpTif: TTiffGraphic;
  EnccryptTif: TTiffGraphic;
  CurFrame: Integer;
begin
  TmpTif := TTiffGraphic.Create;
  Try
    LoadImageToTif(FileName, TmpTif);

    IF IsFileBMP(FileName) Then
    begin
      TmpBMP := TBitmapGraphic.Create;
      Try
        TmpBMP.Assign(TmpTif);
        EncryptDIB(TmpBMP, Key);
        TmpBMP.SaveToFile(FileName);
      Finally
        FreeAndNil(TmpBMP);
      End;
    end
    else IF IsFileTiff(FileName) Then
    begin
      IF TmpTif.FrameCount = 0 Then
      begin
        EncryptDIB(TmpTif, Key);
        TmpTif.SaveToFile(FileName);
      end
      else
      begin
        EnccryptTif := TTiffGraphic.Create;
        Try
          For CurFrame := 1 To TmpTif.FrameCount Do
          begin
            EnccryptTif.Assign(TmpTif.Frames[CurFrame]);
            EncryptDIB(EnccryptTif, Key);
            IF CurFrame = 1 Then
              SaveTifToFile(EnccryptTif, FileName)
            else
              AppendTifToFile(EnccryptTif, FileName,ctColor);
          end;
        Finally
          FreeAndNil(EnccryptTif);
        End;
      end;
    end
    else IF IsFileJpeg(FileName) Then
    begin
      EncryptDIB(TmpTif, Key);
      TmpJPG := TJpegGraphic.Create;
      Try
        TmpJPG.Assign(TmpTif);
        TmpJPG.SaveToFile(FileName);
      Finally
        FreeAndNil(TmpJPG);
      End;
    end
    else
    begin
      EncryptDIB(TmpTif, Key);
      TmpTif.SaveToFile(FileName);
    end;
  Finally
   FreeAndNil(TmpTif);
  End;
end;

procedure Tdm_Image.EncryptDIB(const TmpDib: TDibGraphic;
                                 Key: Integer);
var
  Transform: TEncryptTransform;
begin
  Transform := TEncryptTransform.Create;
  try
   Transform.EncryptionKey := Key;
   Transform.Bidimensional := False;
   Transform.Apply(TmpDib);
  finally
   FreeAndNil(Transform);
  end;
end;

function Tdm_Image.IsOfficeKeyRegistr(Key : WideString) : Boolean;
Var
  Reg : TRegistry;
  CheckPath : WideString;
begin
  Result := False;
  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_CLASSES_ROOT;
    CheckPath := 'CLSID\' + Key;

    if IsOffice64 then
    begin
      if Key = Outlook64AddInRegKey then
      begin
        CheckPath := 'HanibaalOutlook.coHanibaalOutlook\CLSID\';
      end
      else
      if Key = Excel64AddInRegKey then
      begin
        CheckPath := 'HanibaalExcel64.Excel64AddIn\CLSID\';
      end
      else
      if Key = word64AddInRegKey then
      begin
        CheckPath := 'HaniBaalWord64.Word64AddIn\CLSID\';
      end
      else
      if Key = PPT64AddInRegKey then
      begin
        CheckPath := 'HanibaalPPT64.PPT64AddIn\CLSID\';
      end
      else
      if Key = Outlook64AddInRegKey then
      begin
        CheckPath := 'HanibaalOutlook.coHanibaalOutlook\CLSID\';
      end
      else
        CheckPath := 'CLSID\' + Key;
    end;

    try
      if Reg.OpenKeyReadOnly(CheckPath) then
        Result := True;
    except
      Result := False;
    end;
  finally
    Reg.CloseKey;
    FreeAndNil(Reg);
  end;
end;

function  Tdm_Image.IsOfficeUserKeyRegistr(Key : WideString) : Boolean;
Var
  Reg : TRegistry;
begin
  Result := False;
  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    try
      if Reg.OpenKeyReadOnly(Key) Then
        Result := True;
    except
      Result := False;
    end;
  finally
    Reg.CloseKey;
    FreeAndNil(Reg);
  end;
end;

procedure Tdm_Image.FixDibForBarCodeSearch(TmpDib: TBitmapGraphic);
var
  Rslt: Double;
  X: Integer;
  Y, y1: Integer;
  A: TRGB;
  BlackRatio: Double;
  WhiteRatio: Double;
begin
  BlackRatio := { ParamArea.Param_ClearBarCode_Ratio_Black / 100 } 1;
  WhiteRatio := { ParamArea.Param_ClearBarCode_Ratio_White / 100 } 1;

  Rslt := EnDeskewAngleBW(TmpDib, MAX_ANGLE_DEFAULT);
  IF Rslt <> 0 Then
  begin
   DoRotatDIB(TmpDib, Rslt);
  end;

  for X := 0 to TmpDib.Width - 1 do
  begin
   y1 := 0;
   for Y := 0 to TmpDib.Height - 1 do
   begin
     A := TmpDib.Rgb[X, Y];
     if Not((A.Red > 230) and (A.Blue > 230) and (A.Green > 230)) then
     begin
       y1 := y1 + 1;
     end; { if }
   end; { for Y }

   If ((y1 / TmpDib.Height) > BlackRatio) Then
   begin
     for Y := 0 to TmpDib.Height - 1 do
     begin
       A.Red := 0;
       A.Blue := 0;
       A.Green := 0;
       TmpDib.Rgb[X, Y] := A;
     end; { for Y }
   end
   else If ((y1 / TmpDib.Height) <= WhiteRatio) Then
   begin
     for Y := 0 to TmpDib.Height - 1 do
     begin
       A.Red := 255;
       A.Blue := 255;
       A.Green := 255;
       TmpDib.Rgb[X, Y] := A;
     end; { for Y }
   end;

  end;
end;

function  Tdm_Image.GetYearMonthMethode : String;
begin
  Result := 'yyyy-mm';
  // 1 YYYY-MM
  // 2 YYYY_MM
  // 3 YYYY MM
  // 4 Create Month Dir Inside Year Dir

  Case ParamArea.Param_CreateMonthFolderMethode Of
    1 : begin Result := 'yyyy-mm';
        end;
    2 : begin Result := 'yyyy_mm';
        end;
    3 : begin Result := 'yyyy mm';
        end;
    4 : begin Result := 'yyyy\mm';
        end;
  End;
end;

function  Tdm_Image.GetDocumentMafridType_(DocumentDib: TDibGraphic): TMafridType;
begin
  Result := dm_Winsoft.GetDocumentMafridTypeOBR(DocumentDib);
end;

function  Tdm_Image.GetDocumentMafridType_(BarCodesList: TStringList): TMafridType;
begin
  Result := dm_Winsoft.GetDocumentMafridTypeOBR(BarCodesList);
end;

function Tdm_Image.IsImageBarCodeScanSupported(FileName: String): Boolean;
Var
  ExtName: String;
  CurImg: Integer;
begin
  Result := False;
  ExtName := J_ExtFileName(FileName);
  ExtName := Trim(ExtName);
  For CurImg := Low(IMAGE_BARCODE_SCAN_SUPPORTED) To High(IMAGE_BARCODE_SCAN_SUPPORTED) Do
  begin
    IF (UpperCase(IMAGE_BARCODE_SCAN_SUPPORTED[CurImg]) = UpperCase(ExtName)) Then
    begin
      Result := true;
      Break;
    end;
  end;
end;

function  Tdm_Image.GetZXingBarCodesList_(BarCodesList: TStringList;
                                          ClearList: Boolean;
                                          TmpBarCode: TBitmapGraphic): Double;
Var
  CurBC       : Integer;
  BCList      : TStringList;
  CheckStr    : String;
  CurCursor   : TCursor;
  ReadResult  : TReadResult;
  ScanManager : TScanManager;
  TmpBmp      : Graphics.TBitmap; // just to be sure we are really using VCL bitmaps
begin
  If ClearList Then
    BarCodesList.Clear;

  CurCursor := Screen.Cursor;
  Try
    Screen.Cursor := crHourGlass;

    ReadResult  := nil;
    ScanManager := nil;
    TmpBmp := Graphics.TBitmap.Create;

    TmpBmp.Assign(TmpBarCode);
    ScanManager := TScanManager.Create(TBarcodeFormat.Auto, nil);
    ReadResult := ScanManager.Scan(TmpBmp);
    if ReadResult <> nil then
    begin
      if ReadResult.text <> '' Then
        BarCodesList.Add(ReadResult.text);
    end;
  Finally
    FreeAndNil(ScanManager);
    FreeAndNil(ReadResult);
    Screen.Cursor := CurCursor;
  End;
end;

function  Tdm_Image.GetBarCodesList_(BarCodesList: TStringList;
                                      ClearList: Boolean;
                                      TmpBarCode: TBitmapGraphic): Double;
Var
  CurBC     : Integer;
  BCList    : TStringList;
  CheckStr  : String;
  CurCursor : TCursor;
  TmpBmp    : TBitMap;
begin

  If ClearList Then
    BarCodesList.Clear;

  CurCursor := Screen.Cursor;
  Try
    Screen.Cursor := crHourGlass;

    Result := 0;
    if not dm_Protect.LicenseParam.UseDTK then
    begin
      PrmUseBarCodeReaderDTK := False;
      JustWriteToLog('*****     DTK License is missing - dmImage');
    end;

    If PrmUseBarCodeReaderDTK Then
    begin
      JustWriteToLog('USING DTK BAR-CODE READER');
      Result := dmDTK_bc.GetBarCodesListDTK(BarCodesList,
                                             ClearList,
                                             TmpBarCode);
    end;

    //if Result < 1 Then // no barcode found yet
    //begin
    //  JustWriteToLog('USING SD/OBR BAR-CODE READER');
    //  Result := dm_Winsoft.GetBarCodeListOBR(BarCodesList,
    //                                          ClearList,
    //                                          TmpBarCode);
    //end;
    JustWriteToLog('USING OBR BAR-CODE READER');
    Result := dm_Winsoft.GetBarCodeListOBR(BarCodesList,
                                              False,
                                              TmpBarCode);
    IF PrmUseBarCodeReaderSD Then
    begin
      JustWriteToLog('USING SD BAR-CODE READER');
      Result := dm_SDbc.GetBarCodesListSDT(BarCodesList,
                                           False,
                                           TmpBarCode);
    end;

  Finally
    Screen.Cursor := CurCursor;
  End;
end;

function  Tdm_Image.GetBarCodesList_(BarCodesList: TStringList;
                                     ClearList: Boolean;
                                     FileName: WideString): Double;
Var
  TmpTif   : TTiffGraphic;
  TmpBmp   : TBitmapGraphic;
  CurFrame : Integer;
  Rslt     : Double;
begin
  If ClearList Then
    BarCodesList.Clear;

  Result := 0;
  TmpTif := TTiffGraphic.Create;
  TmpBmp := TBitmapGraphic.Create;
  Try
    //LoadImageToTif(FileName,TmpTif);
    LoadImageToTif(FileName, TmpTif, False {Silent}, True {LoadJustFirst}, False {ConvertToBW});
    If Assigned(TmpTif) Then
    begin
      //if TmpTif.FrameCount > 1 Then
      //begin
      //  For CurFrame := 1 To TmpTif.FrameCount do
      //  begin
      //    TmpBmp.Clear;
      //    TmpBmp.Assign(TmpTif.Frames[CurFrame]);
      //    Rslt := GetBarCodesList_(BarCodesList,
      //                             ClearList,
      //                             TmpBmp);
      //    Result := Result + Rslt;
      //  end;
      //end
      //else
      //begin
      //  TmpBmp.Assign(TmpTif);
      //  Result := GetBarCodesList_(BarCodesList,
      //                             ClearList,
      //                             TmpBmp);
      //end;

      TmpBmp.Assign(TmpTif);
      Result := GetBarCodesList_(BarCodesList,
                                   ClearList,
                                   TmpBmp);
    end;
  Finally
    FreeAndNil(TmpBmp);
    FreeAndNil(TmpTif);
  End;
end;

function Tdm_Image.AddImageToPdfFile(SrcFName: String;
                                     DstFName: String;
                                     FitImage : Boolean): Boolean;
Var
  Rslt: longint;
  NewDstFName   : WideString;
  MergeFileName : WideString;
begin
  Result := False;
  IF IsFilePDF(SrcFName) And IsFilePDF(DstFName) Then
  begin
    NewDstFName := RemoveBackSlashChar(ParamArea.WinTempDir) + '\Tmp^' + j_FileName(DstFName);
    Sysutils.DeleteFile(NewDstFName);
    Result := DoMergeFiles(SrcFName,
                            DstFName,
                            NewDstFName,
                            ParamArea.Param_ConvertToPDF_DPI,
                            FitImage);
    if Result then
    begin
      if not Windows.CopyFile(PWideChar(NewDstFName),PWideChar(DstFName),False) then
      begin
        JustWriteToLog('Fail to merge [' + SrcFName + '] to file [' + DstFName + '], Copy Rsult Fail');
      end;
    end
    else
    begin
      JustWriteToLog('Fail to merge [' + SrcFName + '] to file [' + DstFName + ']');
    end;
  end
  else
  begin
    Rslt := AddImgToPdfFile(SrcFName, DstFName, FitImage, 1);
    Result := (Rslt = QuickPdf_NoError);
  end;
end;

Function Tdm_Image.FindPreviewWindows(OpenFileName: String): Boolean;
Var
  CurForm: Integer;
begin
  Result := False;
  LockWindowUpdate(Application.MainForm.Handle);
  Try
    For CurForm := 0 To Application.ComponentCount - 1 Do
    begin
      IF Application.Components[CurForm] is TEnShowImageFrm Then
      begin
        IF UpperCase((Application.Components[CurForm] as TEnShowImageFrm).ImageRec.FileName) = UpperCase(OpenFileName) Then
        begin
          ForceForegroundWindow((Application.Components[CurForm] as TEnShowImageFrm).Handle);
          Result := true;
        end;
      end
      else IF Application.Components[CurForm] is THTMLViewerFrm Then
      begin
        IF UpperCase((Application.Components[CurForm] as THTMLViewerFrm).HtmlFileName) = UpperCase(OpenFileName) Then
        begin
          ForceForegroundWindow((Application.Components[CurForm] as THTMLViewerFrm).Handle);
          Result := true;
        end;
      end;
    end;
  Finally
   LockWindowUpdate(0);
  End;
end;

procedure Tdm_Image.ClosePreviewWindows;
Var
  CurForm: Integer;
begin
  For CurForm := 0 To Application.ComponentCount - 1 Do
  begin
    IF Application.Components[CurForm] is TEnShowImageFrm Then
    begin
      (Application.Components[CurForm] as TEnShowImageFrm).Close;
    end
    else IF Application.Components[CurForm] is THTMLViewerFrm Then
    begin
      (Application.Components[CurForm] as THTMLViewerFrm).Close;
    end
    else
    IF Application.Components[CurForm] is TPDFViewerFrm Then
    begin
      (Application.Components[CurForm] as TPDFViewerFrm).Close;
    end
    else
    IF Application.Components[CurForm] is TOfficeViewerFrm Then
    begin
      (Application.Components[CurForm] as TOfficeViewerFrm).Close;
    end
    else
    IF Application.Components[CurForm] is TGeneralViewerFrm Then
    begin
      (Application.Components[CurForm] as TGeneralViewerFrm).Close;
    end
    else
    IF Application.Components[CurForm] is TDWGViewerFrm Then
    begin
      (Application.Components[CurForm] as TDWGViewerFrm).Close;
    end;
  end;
end;

function Tdm_Image.IsPreviewWindowsExists: Boolean;
Var
  CurForm: Integer;
begin
  Result := False;
  For CurForm := 0 To Application.ComponentCount - 1 Do
  begin
   IF Application.Components[CurForm] is TEnShowImageFrm Then
   begin
    Result := true;
    Break;
   end
   else IF Application.Components[CurForm] is THTMLViewerFrm Then
   begin
    Result := true;
    Break;
   end;
  end;
end;

procedure Tdm_Image.PrintManaSticker(BoxNo: String;
                                      PackageNo: String;
                                      PagesCount: Integer;
                                      UserCrt: String;
                                      ScanStation: String);
begin
end;

procedure Tdm_Image.PrintGraphicDocument(ImageFilename: WideString;
                                   AutoSetOrientation: Boolean);
Begin
  PrintGraphicDocument(ImageFilename,
                       AutoSetOrientation,
                       True,
                       '');
End;

procedure Tdm_Image.PrintGraphicDocument(ImageFilename: WideString;
                                          AutoSetOrientation: Boolean;
                                          DoSelectPrinter: Boolean;
                                          PrintToPrinterName: String);
Var
  Rslt                : TModalResult;
  TifDocumentToPrt    : TTiffGraphic;
  TifFrameToPrt       : TTiffGraphic;
  GraphicPrinter      : TDibGraphicPrinter;
  DftPrinterName      : String;
  OrgPrinterName      : String;
  SelectedPrinterName : String;
  SetPrinterAsDft     : Boolean;
  fPrinterMethode     : TSelectPrintMode;
  fPrintWidth         : Integer;
  fPrintHeight        : Integer;
  fPrintDirection     : Integer;
  CurCursor           : TCursor;
  printCommand        : string;
  printerInfo         : string;
  Device,Driver,Port  : array[0..255] of Char;
  hDeviceMode         : THandle;

  function ReturnEnvisionPrintMode(SelectPrintMode: TSelectPrintMode): TEnvisionPrintMode;
  begin
    Case SelectPrintMode of
     spmOriginalSize:
      Result := EnPrint.pmOriginalSize;
     spmFullPage:
      Result := EnPrint.pmFullPage;
     spmSpecificWidth:
      Result := EnPrint.pmSpecificWidth;
     spmSpecificHeight:
      Result := EnPrint.pmSpecificHeight;
     spmSpecificWidthAndHeight:
      Result := EnPrint.pmSpecificWidthAndHeight;
     spmStretchToPage:
      Result := EnPrint.pmStretchToPage;
     spmSendToPrinter:
      Result := EnPrint.pmOriginalSize;
    else
     Result := EnPrint.pmOriginalSize;
    end;
  end;

  procedure PrintWithAutoPrintJob;
  Var
    CurFrame  : Integer;
    //WidthInc  : Double;
    //HeightInc : Double;
  begin
    SleepForXXXMiliSecend(110, True);
    GraphicPrinter.UsePrintJob := true;
    GraphicPrinter.Title := dm_SiLang.siLangLinked.GetTextOrDefault('IDS_32' (* 'מדפסת גרפית' *) );
    GraphicPrinter.LeftMargin := 0;
    GraphicPrinter.TopMargin := 0;
    GraphicPrinter.PrintMode := ReturnEnvisionPrintMode(fPrinterMethode);

    //WidthInc       := EnMisc.CmToInches(WidthFld.Value/10);
    //HeightInc      := EnMisc.CmToInches(HeightFld.Value/10);

    IF (GraphicPrinter.PrintMode = EnPrint.pmSpecificWidth) Or
      (GraphicPrinter.PrintMode = EnPrint.pmSpecificHeight) Or
      (GraphicPrinter.PrintMode = EnPrint.pmSpecificWidthAndHeight) Then
    begin
     GraphicPrinter.Width := CmToInches(fPrintWidth / 10);
     GraphicPrinter.Height := CmToInches(fPrintHeight / 10);
    end;

    TifDocumentToPrt := TTiffGraphic.Create;
    Try
      dm_Image.LoadImageToTif(ImageFilename, TifDocumentToPrt, False);
      IF TifDocumentToPrt.FrameCount = 0 Then
      begin
        GraphicPrinter.Print(TifDocumentToPrt);
      end
      else
      begin
        TifFrameToPrt := TTiffGraphic.Create;
        Try
          For CurFrame := 1 To TifDocumentToPrt.FrameCount Do
          begin
            TifFrameToPrt.Assign(TifDocumentToPrt.Frames[CurFrame]);
            GraphicPrinter.Print(TifFrameToPrt);
          end;
        Finally
          FreeAndNil(TifFrameToPrt);
        End;
      end;
    Finally
     FreeAndNil(TifDocumentToPrt);
    End;
  end;

begin
  inherited;
  IF Not GetDftPrinterName(DftPrinterName) Then
   Raise EAppError.Create(Format(dm_SiLang.siLangLinked.GetTextOrDefault('IDS_33' (* 'מדפסת ברירת מחדל לא נמצאה' *) ),
     [ImageFilename]));

  OrgPrinterName := DftPrinterName;
  IF Trim(PrintToPrinterName) = '' Then
   DoSelectPrinter := true;

  IF Not DoSelectPrinter Then
  begin
   IF GetPrinterIndex(PrintToPrinterName) = ITEM_NOT_FOUND Then
     DoSelectPrinter := true;
  end;

  IF DoSelectPrinter Then
  begin
    JustWriteToLog('Print graphic file to selected printer');
    SelectuserPrinterFrm := TSelectuserPrinterFrm.Create(Self);
    try
      SelectuserPrinterFrm.CanPrinterAsDefault := true;
      SelectuserPrinterFrm.FileToPrint := ImageFilename;
      Rslt := SelectuserPrinterFrm.ShowModal;
      SelectedPrinterName := SelectuserPrinterFrm.ChossenPrinterName;
      SetPrinterAsDft     := SelectuserPrinterFrm.SetPrinterAsDefault;
      fPrinterMethode     := SelectuserPrinterFrm.fPrinterMethode;
      fPrintWidth         := SelectuserPrinterFrm.fPrintWidth;
      fPrintHeight        := SelectuserPrinterFrm.fPrintHeight;
      fPrintDirection     := SelectuserPrinterFrm.fPrintDirection;

      PrintToPrinterName  := SelectedPrinterName;
    finally
      FreeAndNil(SelectuserPrinterFrm);
    end;

    IF Rslt <> mrOk then
      Exit;
  end
  else
  begin
    //yossi change for club-hotel fPrinterMethode := spmSendToPrinter;
    JustWriteToLog('Print graphic file to Graphic printer, Using Original Size');
    fPrinterMethode := spmOriginalSize;
    SelectedPrinterName := PrintToPrinterName;
    SetPrinterAsDft := False;
  end;

  // if DoSelectPrinter and
  // rslt = mrOK
  IF fPrinterMethode = spmSendToPrinter Then
  begin
    IF Trim(PrintToPrinterName) = '' Then
      PrintToPrinterName := DftPrinterName;
    IF UpperCase(PrintToPrinterName) = UpperCase(DftPrinterName) Then
    begin
      printCommand := 'print';
      printerInfo  := '';

      Try
        //LockWindowUpdate(Application.Handle);
        SleepForXXXMiliSecend(110, True);
        ShellExecute(Application.Handle, PChar(printCommand), PChar(ImageFilename), PChar(printerInfo), nil, SW_SHOW) ;
      Finally
        //LockWindowUpdate(0);
        SleepForXXXMiliSecend(110, True);
      End;
    end
    else
    begin
      // the printto is not working well
      // set printer to dft, print, return default to orgiginal dft

      //printCommand := 'printto';
      //Printer.PrinterIndex := GetPrinterIndex(PrintToPrinterName);
      //Printer.GetPrinter(Device, Driver, Port, hDeviceMode) ;
      //printerInfo := Format('"%s" "%s" "%s"', [Device, Driver, Port]) ;
      Try
        SetDefaultPrinter2(PrintToPrinterName);
        Try
          printCommand := 'print';
          printerInfo  := '';

          //LockWindowUpdate(Application.Handle);
          SleepForXXXMiliSecend(110, True);
          ShellExecute(Application.Handle, PChar(printCommand), PChar(ImageFilename), PChar(printerInfo), nil, SW_SHOW) ;
        Finally
          //LockWindowUpdate(0);
          SleepForXXXMiliSecend(110, True);
        End;
      Finally
        SetDefaultPrinter2(OrgPrinterName);
        SleepForXXXMiliSecend(110, True);
      End;
    end;
  end
  else
  begin
    GetDftPrinterName(DftPrinterName);
    OrgPrinterName := DftPrinterName;
    Try
      IF UpperCase(SelectedPrinterName) <> UpperCase(OrgPrinterName) Then
      begin
        SetDefaultPrinter2(SelectedPrinterName);
        Printer.PrinterIndex := Printer.Printers.IndexOf(SelectedPrinterName);
        SleepForXXXMiliSecend(110, True);
      end;

      GraphicPrinter := TDibGraphicPrinter.Create;
      try
        CurCursor := Screen.Cursor;
        Screen.Cursor := crHourGlass;
        PrintWithAutoPrintJob;
      finally
        FreeAndNil(GraphicPrinter);
        Screen.Cursor := CurCursor;
      end;
    Finally
      IF UpperCase(SelectedPrinterName) <> UpperCase(OrgPrinterName) Then
      begin
        SetDefaultPrinter2(OrgPrinterName);
        Printer.PrinterIndex := Printer.Printers.IndexOf(OrgPrinterName);
        SleepForXXXMiliSecend(110, True);
      end;
    End;
  end;
end;

procedure Tdm_Image.JustPrintFile(Filename: WideString;
                                   UseDialog : Boolean;
                                   DoWait : Boolean = False);
begin
  if UseDialog then
  begin
    if Not PrintDialog.Execute then
      Exit;
  end;

  RunWinExe(Application.Handle, // Handle
            soPrint,            // Operation
            '',                 // Dft Directory
            True,               // Do Hide
            DoWait,             // Do Wait
            Filename,           // Command
            '');                // Param01..10
end;

function Tdm_Image.QP_PrintPDF(FileName : WideString) : Integer;
Var
  PDFLibrary   : TQuickPDF;
  UnlockResult : Integer;
  dlgPrint     : TPrintDialog;
  PrinterName  : String;
  PrintOption  : Integer;
begin
  Result := -99;
  if Not FileExists(FileName) then
    Exit;

  PDFLibrary := TQuickPDF.Create;
  Try
    UnlockResult := PDFLibrary.UnlockKey(edtLicenseKey);
    if UnlockResult <> 1 then
      Exit;
    if PDFLibrary.LoadFromFile(FileName, '') <> 1 then
    begin
      MessageDlg('The PDF could not be opened.', mtError, [mbOK], 0);
      Exit;
    end;

    dlgPrint := TPrintDialog.Create(nil);
    Try
      dlgPrint.Options    := [poPageNums , poSelection, poPrintToFile, poWarning];
      dlgPrint.PrintRange := prSelection;
      dlgPrint.MinPage    := 1;
      dlgPrint.MaxPage    := PDFLibrary.PageCount;
      dlgPrint.FromPage   := 1;
      dlgPrint.ToPage     := PDFLibrary.PageCount;
      if dlgPrint.Execute then
      begin
        PrinterName  := Printer.Printers[Printer.PrinterIndex];
        PrintOption := PDFLibrary.PrintOptions(1, 1, dm_SiLang.siLangLinked.GetTextOrDefault('IDS_0' (* 'מערכת ניהול מסמכים' *) ));
        Result := PDFLibrary.PrintDocument(PrinterName,
                                           dlgPrint.FromPage, dlgPrint.ToPage,
                                           PrintOption);
      end;
    Finally
      FreeAndNil(dlgPrint);
    End;
  Finally
    FreeAndNil(PDFLibrary);
  End;
end;

procedure Tdm_Image.PrintAnyDocument(ImageFilename   : WideString;
                                     DoSelectPrinter : Boolean);
Var
  MaxPages   : Integer;
  PrinterDftName : String;
  PrinterTo      : string;
  PrintCommand   : string;
  PrinterInfo    : string;
  Device, Driver, Port: array[0..255] of Char;
  hDeviceMode: THandle;
  //no need - take printer global var - Printer : TPrinter;

begin
  IF dm_Image.IsFileCanBePrinted(ImageFilename) Then
  begin
    // first set the dst printer
    if not GetDftPrinterName(PrinterTo) Then
      PrinterTo := '';
    if DoSelectPrinter then
    begin
      MaxPages := 999;
      PrintDialog.ToPage  := MaxPages;
      PrintDialog.MaxPage := MaxPages;
      If not GetDftPrinterName(PrinterDftName) Then
        PrinterDftName := '';
      if PrintDialog.Execute then
      begin
        //Printer := TPrinter.Create;
        Try
           PrinterTo := GetPrinterByIndex(Printer.PrinterIndex);
        Finally
          //Printer.Free;
        End;
      end;
    end;

    PrintCommand := 'print';
    PrinterInfo := '';
    if PrinterTo <> PrinterDftName Then
    begin
      printCommand := 'printto';
      Printer.PrinterIndex := Printer.PrinterIndex;
      Printer.GetPrinter(Device, Driver, Port, hDeviceMode) ;
      printerInfo := Format('"%s" "%s" "%s"', [Device, Driver, Port]) ;
    end;

    //ShellExecute(Application.Handle, 'print', PWideChar(ImageFilename), nil, nil, SW_HIDE) ;
    Try
      LockWindowUpdate(Application.Handle);
      ShellExecute(Application.Handle, PChar(PrintCommand), PChar(ImageFilename), PChar(PrinterInfo), nil, SW_HIDE) ;
    Finally
      LockWindowUpdate(0);
    End;
  end
  else
  begin
    ShowImage_External(ImageFilename);
  end;
end;

function  Tdm_Image.GetTextFromDocumentFile(DocFileName : WideString) : WideString;
begin
  Result := '';
  Try
    IF dm_Image.IsFileDOC(DocFileName) Or
       dm_Image.IsFileDOCX(DocFileName) Then
    begin
      Result := GetTextFromWordFile(DocFileName);
    end
    else
    IF dm_Image.IsFileExcel(DocFileName) And
       ParamArea.Param_GetXlsOcrWords Then
    begin
      Result := GetWordsListFromExcel(DocFileName);
    end
    else
    IF dm_Image.IsFilePDF(DocFileName) Then
    begin
      Result := GetTextFromPDFFile(DocFileName);
    end
    else
    IF dm_Image.IsFileOutLookMsg(DocFileName) Then
    begin
      Result := GetTextFromMSGFile(DocFileName);
    end;
  Except;
    Result := '';
  End;

  if ParamArea.Param_OnOcr_RemoveDuplecate then
    Result := dm_Ocr.RemoveDuplicateWords(Result);
  if ParamArea.Param_OnOcr_DropSmallWord then
    Result := dm_Ocr.RemoveSmallWords(Result);

  Result := Trim(Result);
end;

function  Tdm_Image.GetTextFromWordFile(WordFileName : WideString) : WideString;
Var
  MsgStr : String;
begin
  Result := '';
  IF ParamArea.Param_WriteToLog Then
  begin
    MsgStr := 'function  Tdm_Image.GetTextFromWordFile-WordFileName='+WordFileName;
    dm_Main.MzWriteLog(MsgStr,MzLog_FromProcess);
  end;

  Try
    Result := GetWordsListFromWord(WordFileName);
  Except;
    Result := '';
  End;

  IF ParamArea.Param_WriteToLog Then
  begin
    MsgStr := 'Result of Word Pasing='+Result;
    dm_Main.MzWriteLog(Trim(MsgStr),MzLog_FromProcess);
  end;
end;

function  Tdm_Image.GetTextFromPDFFile(PdfFileName : WideString) : WideString;
Var
  QuickPdf     : TQuickPDF;
  CurPage      : Integer;
  PosFound     : Integer;
  PageText     : WideString;
  UnlockResult : Integer;
  MyWordList   : TStringList;
  fKeyWord     : WideString;
  CurItem      : Integer;
  TickTime1,
  TickTime2 : LongInt;
  Str          : String;
begin
  Result := '';
  MyWordList := TStringList.Create;
  QuickPdf   := TQuickPdf.Create;
  Try
    UnlockResult := QuickPdf.UnlockKey(edtLicenseKey);
    if UnlockResult <> 1 then
      Exit;
    TickTime1 := GetTickCount;
    Try
      IF QuickPdf.LoadFromFile(PdfFileName,'') = 1 Then
      begin
        For CurPage := 1 To QuickPdf.PageCount Do
        begin
          PageText := QuickPdf.ExtractFilePageText(PdfFileName,'', CurPage, 7);
          MyWordList.Clear;

          PageText := RemoveAllEnterChar(PageText);
          PageText := RemoveAllTabChar(PageText);
          CutStringToWords(PageText, MyWordList , True);

          For CurItem := 0 to MyWordList.Count - 1 do
          begin
            Str := Trim(MyWordList.Strings[CurItem]);
            if Str <> '' then
            begin
              Str := FlipOnlyHebChar(Str);
              MyWordList.Strings[CurItem] := Str;
            end;
          end;

          Result := Result + ' ' + MyWordList.Text;
        end;
      end;
    Except;
      Result := '';
    End;
    TickTime2 := GetTickCount;
    JustWriteToLog('TickTime =' + FloatToStr((TickTime2-TickTime1)/1000));;
  Finally
    FreeAndNil(MyWordList);
    FreeAndNil(QuickPdf);
  End;
end;

function  Tdm_Image.GetTextFromMSGFile(MsgFileName : WideString) : WideString;
Var
  RunExeName     : WideString;
  ResultFileName : WideString;
  FileToLoad     : WideString;
  Success        : Boolean;
  FileLists      : TStringList;
  CurTime        : Integer;
  CurFile        : Integer;
  tt0, tt1, tt2  : Integer;
  MS             : TMemoryStream;
  CurCursor      : TCursor;
begin
  Result := '';

  if IsOffice64 then
  begin
    RunExeName := RemoveBackSlashChar(j_PathFileName(Application.ExeName)) + '\' + mzGetMsgBody64;
  end
  else
  begin
    RunExeName := RemoveBackSlashChar(j_PathFileName(Application.ExeName)) + '\' + mzGetMsgBody32;
  end;

  if FileExists(RunExeName) then
  begin
    ResultFileName := RemoveBackSlashChar(ParamArea.WinTempDir) + '\TmpResult' + IntToStr(GetTickCount) + '.Txt';
    windows.DeleteFile(PWideChar(ResultFileName));

    RunWinExe(Application.Handle, // Handle
                soOpen,            // Operation
                j_PathFileName(Application.ExeName), // Dft Directory
                True,               // Do Hide
                False,              // Do Wait
                RunExeName,         // Command
                MsgFileName,        // Param01
                ResultFileName,     // Param02
                AppEncryPass);      // Param03

    tt1 := GetTickCount;
    Success := False;
    CurCursor := Screen.Cursor;
    Try
      Screen.Cursor := crHourGlass;
      For CurTime := 0 To Max_Try_LoadMsgBodyCount Do
      begin
        SleepForXXXMiliSecend(500,True);
        if FileExists(ResultFileName) then
        begin
          Success := True;
          Break;
        end;
      end;
    Finally
      Screen.Cursor := CurCursor;
    End;

    tt2 := GetTickCount;
    tt0 := tt2-tt1;
    JustWriteToLog('Load Messae Body took ' + FloatToStr(tt0/1000) + dm_SiLang.siLangLinked.GetTextOrDefault('IDS_34' (* ' שניות' *) ));

    If Success Then
    begin
      FileLists := TStringList.Create;
      Try
        FileLists.LoadFromFile(ResultFileName);
        For CurFile := 0 to FileLists.Count - 1 do
        begin
          FileToLoad := Trim(FileLists.Strings[CurFile]);
          JustWriteToLog('Get Message Text - Test File Name =  ' + FileToLoad);
          if FileExists(FileToLoad) And
             IsFileTxt(FileToLoad) then
          begin
            MS := TMemoryStream.Create;
            Try
              MS.LoadFromFile(FileToLoad);
              MS.Position := 0;
              SetString(Result, PWideChar(MS.Memory), MS.Size div SizeOf(Char));
            Finally
              MS := nil;
            End;
            Break;
          end;
        end;
      Finally
        FreeAndNil(FileLists);
      End;
    end
  end;
end;

procedure Tdm_Image.SetTextFromMSGToDb(MzIdent : Integer);
Var
  RunExeName : WideString;
  ResultFileName : WideString;
  FileToLoad : WideString;
  Success : Boolean;
  FileLists : TStringList;
  CurTime : Integer;
  CurFile : Integer;
  tt0, tt1, tt2 : Integer;
  MS : TMemoryStream;
begin
  if IsOffice64 then
  begin
    RunExeName := RemoveBackSlashChar(j_PathFileName(Application.ExeName)) + '\' + mzSetMsgBodyToDB64;
  end
  else
  begin
    RunExeName := RemoveBackSlashChar(j_PathFileName(Application.ExeName)) + '\' + mzSetMsgBodyToDB32;
  end;

  if FileExists(RunExeName) then
  begin
    RunWinExe(Application.Handle, // Handle
                soOpen,             // Operation
                j_PathFileName(Application.ExeName), // Dft Directory
                True,               // Do Hide
                False,              // Do Wait
                RunExeName,         // Command
                '');                // Param01..10
  end;
end;

procedure Tdm_Image.SetJpgIconToDb(MzDocID : Integer);
Var
  RunExeName : WideString;
  ResultFileName : WideString;
  FileToLoad : WideString;
  Success : Boolean;
  FileLists : TStringList;
  CurTime : Integer;
  CurFile : Integer;
  tt0, tt1, tt2 : Integer;
  MS : TMemoryStream;
begin
  //if IsWindows64 then
  //begin
  //  RunExeName := RemoveBackSlashChar(j_PathFileName(Application.ExeName)) + '\' + mzSetJpgIconToDB64;
  //end
  //else
  //begin
  //  RunExeName := RemoveBackSlashChar(j_PathFileName(Application.ExeName)) + '\' + mzSetJpgIconToDB32;
  //end;

  RunExeName := mzSetJpgIconToDB;

  JustWriteToLog('Try Create Jpg Icon fo the First Page, DocID = ' + IntToStr(MzDocID));
  if FileExists(RunExeName) then
  begin
    RunWinExe(Application.Handle, // Handle
              soOpen,             // Operation
              j_PathFileName(Application.ExeName), // Dft Directory
              True,               // Do Hide
              False,              // Do Wait
              RunExeName,         // Command
              IntToStr(MzDocID)); // Param01..10
  end;
end;

function Tdm_Image.ReplaceWordDoc_Dchiya(DataToReplace : TWordDoc_Dchiya;
                                          DocFileName : String;
                                          SaveToFileName : String) : Boolean;
Var
  MyWrdApp: TWordApplication;
  MyWrdDoc: word2000.TWordDocument;
  WFormat,
   WFileName : OleVariant;

  FindText,
    Wrap,
    oleTrue,
    ReplaceWith,
    Replace : OleVariant;
  TmpFile   : String;
  SaveToDir : String;
  SaveChanges: OleVariant;
Const
  wdDoNotSaveChanges = 0;
  wdFindContinue = 1;
  wdReplaceOne = 1;
  wdReplaceAll = 2;

  procedure FindReplace(inSearch, InReplace : String);
  begin
    Try
      if (Length(InReplace) > 255) then
      begin
        InReplace := Copy(InReplace, 1, 255);
      end;

      FindText    := inSearch;
      ReplaceWith := InReplace;
      Wrap        := wdFindContinue;
      Replace     := wdReplaceAll;

      MyWrdApp.Selection.Find.Execute(FindText,
                                    EmptyParam,  // MatchCase: OleVariant;
                                    EmptyParam,  // MatchWholeWord: OleVariant;
                                    EmptyParam,  // MatchWildcards: OleVariant;
                                    EmptyParam,  // MatchSoundsLike: OleVariant;
                                    EmptyParam,  // MatchAllWordForms: OleVariant;
                                    EmptyParam,  // Forward: OleVariant;
                                    Wrap,        // Wrap: OleVariant;
                                    EmptyParam,  // Format: OleVariant;
                                    ReplaceWith, // ReplaceWith: OleVariant;
                                    Replace,     // MatchControl: OleVariant;
                                    EmptyParam,
                                    EmptyParam,
                                    EmptyParam,
                                    EmptyParam);
    Except
      on E : Exception do
        begin
          Raise Exception.Create(E.Message);
        end;
    End;
  end;

  function OpenDoc( Word : TWordApplication; FName : WideString ) : _Document;
  var
    FileName, ConfirmConversions, ReadOnly, AddToRecentFiles,
    Revert, Format, Visible: OLEVariant;
  begin
    FileName           := FName;
    ConfirmConversions := False;
    ReadOnly           := False;
    AddToRecentFiles   := False;
    Revert             := True;
    Format             := wdOpenFormatAuto;
    Visible            := True;

    Result := Word.Documents.Open( FileName, ConfirmConversions, ReadOnly,
      AddToRecentFiles, EmptyParam, EmptyParam, Revert, EmptyParam,
      EmptyParam, Format, EmptyParam, Visible{, EmptyParam, EmptyParam, EmptyParam} );
  end;

begin
  Result := False;
  IF Not FileExists(DocFileName) Then
    Exit;

  SaveToDir := j_PathFileName(SaveToFileName);
  IF Not DirectoryExists(SaveToDir) Then
    ForceDirectories(SaveToDir);
  IF Not DirectoryExists(SaveToDir) Then
    Exit;
  TmpFile := RemoveBackSlashChar(ParamArea.WinTempDir) + '\TmpDoc.' + J_ExtFileName(DocFileName);
  IF FileExists(TmpFile) Then
    DeleteFile(TmpFile);

  CopyFile(Pchar(DocFileName),Pchar(TmpFile),False);
  IF Not FileExists(TmpFile) Then
    Exit;

  MyWrdApp:= TWordApplication.Create(nil);
  MyWrdDoc:= word2000.TWordDocument.Create(nil);
  Try
    MyWrdApp.ConnectKind := ckRunningOrNew;
    MyWrdApp.Connect;
    try
      MyWrdDoc.ConnectKind :=ckAttachToInterface;
      MyWrdDoc.ConnectTo(OpenDoc(MyWrdApp, TmpFile));

      FindReplace('X00001', FormatDateTime('dd/mm/yyyy',Now));  //
      FindReplace('X00002', DataToReplace.OppCompanyName);      // שם חברה נגדית
      FindReplace('X00003', FormatDateTime('dd/mm/yyyy',DataToReplace.NezekDate));  // תאריך נזק
      FindReplace('X00004', DataToReplace.OppCarNo);            // רכב נגדי
      FindReplace('X00005', DataToReplace.CarNo);               // רכבנו
      FindReplace('X00006', DataToReplace.Text);                // סיבת דחייה

      WFormat   := wdFormatDocument;
      WFileName := OLEVariant(SaveToFileName);
      oleTrue   := True;

      MyWrdApp.ActiveDocument.SaveAs(WFileName
                                      ,WFormat
                                      ,EmptyParam
                                      ,EmptyParam
                                      ,EmptyParam
                                      ,EmptyParam
                                      ,EmptyParam
                                      ,EmptyParam
                                      ,EmptyParam
                                      ,EmptyParam
                                      ,EmptyParam
                                      {,EmptyParam
                                      ,EmptyParam
                                      ,EmptyParam
                                      ,EmptyParam
                                      ,EmptyParam});
      SaveChanges := wdDoNotSaveChanges;
      MyWrdApp.ActiveDocument.Close(SaveChanges,EmptyParam,EmptyParam);

      Sleep(10);
      Application.ProcessMessages;
      MyWrdApp.Quit;
    finally
    end;

    IF FileExists(SaveToFileName) Then
      Result := True;
  finally
    MyWrdDoc := nil;
    MyWrdApp := nil;

    Sleep(10);
    Application.ProcessMessages;
  end;
end;

end.





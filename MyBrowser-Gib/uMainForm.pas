unit uMainForm;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs, FMX.StdCtrls,
  FMX.Controls.Presentation, FMX.Layouts, FMX.WebBrowser, FMX.Edit, FMX.Objects;

type
  TForm1 = class(TForm)
    LayoutUrl: TLayout;
    edURL: TEdit;
    LayoutBrowser: TLayout;
    WebBrowser: TWebBrowser;
    LayoutTop: TLayout;
    BtnBack: TButton;
    BtnForward: TButton;
    BtnRefresh: TButton;
    Label1: TLabel;
    Layout1: TLayout;
    MyAniIndicator: TAniIndicator;
    Image1: TImage;
    imgClean: TImage;
    imgYnet: TImage;
    imgGoogle: TImage;
    imgClaude: TImage;
    imgGemini: TImage;
    imgGo: TImage;
    procedure FormCreate(Sender: TObject);
    procedure BtnBackClick(Sender: TObject);
    procedure BtnForwardClick(Sender: TObject);
    procedure BtnRefreshClick(Sender: TObject);
    procedure WebBrowserDidStartLoad(ASender: TObject);
    procedure WebBrowserDidFinishLoad(ASender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormDeactivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Image1Click(Sender: TObject);
    procedure imgCleanClick(Sender: TObject);
    procedure imgYnetClick(Sender: TObject);
    procedure imgGoogleClick(Sender: TObject);
    procedure imgClaudeClick(Sender: TObject);
    procedure imgGeminiClick(Sender: TObject);
    procedure imgGoClick(Sender: TObject);
  private
    { Private declarations }
    procedure ClearAllHistory;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation
{$R *.fmx}
{$R *.LgXhdpiPh.fmx ANDROID}
{$R *.NmXhdpiPh.fmx ANDROID}
{$R *.iPhone47in.fmx IOS}

uses
  Androidapi.JNI.WebKit,
  Androidapi.Helpers,
  Androidapi.JNIBridge;

procedure TForm1.BtnBackClick(Sender: TObject);
begin
  WebBrowser.GoBack;
end;

procedure TForm1.ClearAllHistory;
var
  LocalObj: ILocalObject;
  WebView: JWebView;
begin
  // Clear all cookies
  TJCookieManager.JavaClass.getInstance.removeAllCookies(nil);
  TJCookieManager.JavaClass.getInstance.flush;

  if Supports(WebBrowser, ILocalObject, LocalObj) then
  begin
    try
      WebView := TJWebView.Wrap(LocalObj.GetObjectID);
      if WebView <> nil then
      begin
        WebView.clearHistory;
        WebView.clearCache(True);
        WebView.clearFormData;
      end;
    except
      // Silently ignore if the native WebView cannot be accessed
    end;
  end;
end;

procedure TForm1.BtnForwardClick(Sender: TObject);
begin
  WebBrowser.GoForward;
end;

procedure TForm1.BtnRefreshClick(Sender: TObject);
begin
  WebBrowser.Reload;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  ClearAllHistory;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  edURL.Text := '';
  MyAniIndicator.Parent := LayoutUrl;
  MyAniIndicator.Align  := TAlignLayout.Right;
  MyAniIndicator.Enabled := False;
  MyAniIndicator.Visible := False;
  MyAniIndicator.BringToFront;
end;

procedure TForm1.FormDeactivate(Sender: TObject);
begin
  ClearAllHistory;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  ClearAllHistory;
end;

procedure TForm1.Image1Click(Sender: TObject);
begin
  Application.Terminate;
end;

procedure TForm1.imgGeminiClick(Sender: TObject);
begin
  edURL.Text := 'https://gemini.google.com/app';
  imgGoClick(Sender);
end;

procedure TForm1.imgGoClick(Sender: TObject);
begin
  if Trim(edURL.Text) <> '' then
    WebBrowser.URL := Trim(edURL.Text);
end;

procedure TForm1.imgGoogleClick(Sender: TObject);
begin
  edURL.Text := 'https://google.com';
  imgGoClick(Sender);
end;

procedure TForm1.imgClaudeClick(Sender: TObject);
begin
  edURL.Text := 'https://claude.ai/new';
  imgGoClick(Sender);
end;

procedure TForm1.imgCleanClick(Sender: TObject);
begin
  Try
    MyAniIndicator.Visible := True;
    MyAniIndicator.Enabled := True;
    Sleep(1000);
    ClearAllHistory;
  Finally
    MyAniIndicator.Enabled := False;
    MyAniIndicator.Visible := False;
  End;
end;

procedure TForm1.imgYnetClick(Sender: TObject);
begin
  edURL.Text := 'https://ynet.co.il';
  imgGoClick(Sender);
end;

procedure TForm1.WebBrowserDidFinishLoad(ASender: TObject);
Var
  NewUrl : String;
begin
  MyAniIndicator.Enabled := False;
  MyAniIndicator.Visible := False;
  NewUrl := WebBrowser.URL;
  if not SameText(edURL.Text, WebBrowser.URL) then
    edURL.Text := WebBrowser.URL;
end;

procedure TForm1.WebBrowserDidStartLoad(ASender: TObject);
begin
  MyAniIndicator.Visible := True;
  MyAniIndicator.Enabled := True;
  MyAniIndicator.BringToFront;
end;

end.

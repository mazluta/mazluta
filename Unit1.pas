unit Unit1;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs, FMX.StdCtrls,
  FMX.Controls.Presentation, FMX.Layouts, FMX.WebBrowser, FMX.Edit;

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
    btnClear: TButton;
    btGo: TButton;
    procedure FormCreate(Sender: TObject);
    procedure BtnBackClick(Sender: TObject);
    procedure BtnForwardClick(Sender: TObject);
    procedure BtnRefreshClick(Sender: TObject);
    procedure btnClearClick(Sender: TObject);
    procedure btGoClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.fmx}
{$R *.LgXhdpiPh.fmx ANDROID}

uses
  Androidapi.JNI.WebKit,
  Androidapi.Helpers,
  Androidapi.JNIBridge;

procedure TForm1.btGoClick(Sender: TObject);
begin
  if Trim(edURL.Text) <> '' then
    WebBrowser.URL := Trim(edURL.Text);
end;

procedure TForm1.BtnBackClick(Sender: TObject);
begin
  WebBrowser.GoBack;
end;

procedure TForm1.btnClearClick(Sender: TObject);
var
  WebView: JWebView;
  LocalObj: ILocalObject;
begin
  // Clear all cookies
  TJCookieManager.JavaClass.getInstance.removeAllCookies(nil);
  TJCookieManager.JavaClass.getInstance.flush;

  // Access the native WebView and clear history/cache/form data
  // FIX: guard with nil check and exception handler to avoid EJNIException (class 10)
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

procedure TForm1.FormCreate(Sender: TObject);
begin
  // FIX: removed WebBrowser.CleanupInstance — that method destroys managed fields
  // of a live object (strings/interfaces/dynamic arrays), corrupting TWebBrowser
  // internals immediately and causing "exception class 10" on Android.
  edURL.Text := '';
end;

end.

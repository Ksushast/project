unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComObj, StdCtrls;

type
  TForm1 = class(TForm)
    GroupBox1: TGroupBox;
    Button1: TButton;
    Button2: TButton;
    Button4: TButton;
    GroupBox2: TGroupBox;
    Button3: TButton;
    Button5: TButton;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  w: Variant;
  e: Variant;

const
  ExcelApp = 'Excel.Application';

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
begin
 w:=CreateOleObject('Word.Application');
  w.Documents.Add;
 w.Visible:=true;

end;

procedure TForm1.Button2Click(Sender: TObject);
const
  wdAlignParagraphCenter = 1;
  wdAlignParagraphLeft = 0;
  wdAlignParagraphRight = 2;
var
  wdApp, wDoc, wR : Variant;
begin
w:=CreateOleObject('Word.Application');
wDoc:=w.Documents.Open(ExtractFilePath(paramstr(0))+'/MSWord.doc');
w.Visible:=true;
 w.ScreenUpdating := False;
  try
    wR:= wDoc.Range;
    wR.InsertBefore('text1.'#13#10);
    wR.Font.Name := 'Times New Roman';
    wR.Font.Italic := True;
    wR.Font.Size := 14;
    wR.ParagraphFormat.Alignment := wdAlignParagraphLeft;
    wR.InsertAfter(#13#10);

  finally
    w.ScreenUpdating := True;
    end;
end;

procedure TForm1.Button4Click(Sender: TObject);
begin
 w.ActiveDocument.Close(True);
 w.Quit;
end;

procedure TForm1.Button3Click(Sender: TObject);
var
  i,j: byte;
begin
   e:=CreateOleObject('Excel.Application');
   e.WorkBooks.Add;
  for i:=1 to 15 do
  e.Cells[i,1].value:=i;
  e.Visible:=true;
end;

procedure TForm1.Button5Click(Sender: TObject);
begin
  e.ActiveWorkbook.SaveAs('E:\excel.xls');
  e.Quit;
  close;
end;

end.



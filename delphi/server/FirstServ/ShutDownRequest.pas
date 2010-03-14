unit ShutDownRequest;

interface

uses Windows,SysUtils,Classes,Graphics,Forms,Controls,StdCtrls,Buttons,ExtCtrls;

type
  TShutDownDlg = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    RadioGroup1: TRadioGroup;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ShutDownDlg: TShutDownDlg;

implementation

{$R *.DFM}

end.

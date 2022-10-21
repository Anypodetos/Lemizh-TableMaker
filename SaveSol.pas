unit SaveSol;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, ComCtrls, Gauges;

type
  TSaveSolv = class(TForm)
    Label1: TLabel;
    Image1: TImage;
    Panel: TPanel;
    Gauge: TGauge;
  private
  public
  end;

var
  SaveSolv: TSaveSolv;

implementation

{$R *.DFM}

end.

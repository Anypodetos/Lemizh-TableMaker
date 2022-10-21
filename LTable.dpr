program LTable;

uses
  Forms,
  LTableMaker in 'LTableMaker.pas' {TableMaker},
  Opts in 'Opts.pas' {Options},
  SaveSol in 'SaveSol.pas' {SaveSolv},
  Random in 'Random.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.Title := 'Lemizh Table Maker';
  Application.CreateForm(TTableMaker, TableMaker);
  Application.CreateForm(TOptions, Options);
  Application.CreateForm(TSaveSolv, SaveSolv);
  Application.Run;
end.

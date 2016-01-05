program wordPrint;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  Classes,
  ActiveX,
  WordFunctions in 'src\WordFunctions.pas';

var
  i : integer;
  paramlist : TStringlist;

  LogResult : String;
begin

  paramlist := TStringlist.create;
  CoInitialize(nil);
    try
      for i := 1 to ParamCount do
      begin
       paramlist.Add(ParamStr(i));
      end;
    finally

    end;

    OpenandPrint(ParamList[0],ParamList[1]);

     writeln(LogResult);
end.


(*
C:\Users\toby\Documents\GitHub\wordprint>wordprint "http://backoffice.cookingisfun.ie/reports/dailysheet_week.php?DailySheetID=731&authtoken=3c5a24bc-90e1-4fa5-8b88-706fa0f1bb33" "Canon iR-ADV 6055/6065 UFR II"
















*)
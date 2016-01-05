unit WordFunctions;

interface
uses Classes,   ActiveX, ComObj, WinINet, Variants, sysutils, Types;

  function OpenandPrint(FileorURL, Printer : String): String;

implementation

function OpenandPrint(FileorURL, Printer : String): String;
var WordApp, ActiveDoc : Olevariant;
marginvalue : real;
begin

    Wordapp :=  CreateOleObject('Word.Application');
    Wordapp.Visible := false;




    ActiveDoc := WordApp.Documents.Open(FileorURL);
        ActiveDoc.PageSetup.Orientation := 1; //landscape
    MarginValue := 0.6 * 28;

    ActiveDoc.PageSetup.TopMargin := MarginValue;
        ActiveDoc.PageSetup.BottomMargin := MarginValue;
        ActiveDoc.PageSetup.LeftMargin := MarginValue ;
        ActiveDoc.PageSetup.RightMargin := MarginValue ;


    WordApp.ActivePrinter :=  Printer; //'HP LaserJet 400 M401dne UPD PCL 6'; // '\\SERVER\Canon iR-ADV 6055/6065 UFR II';

    WordApp.PrintOut();

    //C:\Users\toby\Documents\GitHub\wordprint>wordprint "http://backoffice.cookingisfun.ie/reports/dailysheet_week.php?DailySheetID=731&authtoken=3c5a24bc-90e1-4fa5-8b88-706fa0f1bb33"






 (*       Documents.Open FileName:="http://www.google.com/", ConfirmConversions:= _
        False, ReadOnly:=True, AddToRecentFiles:=False, PasswordDocument:="", _
        PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", _
        WritePasswordTemplate:="", Format:=wdOpenFormatAuto, XMLTransform:=""
    Application.PrintOut FileName:="", Range:=wdPrintRangeOfPages, Item:= _
        wdPrintDocumentContent, Copies:=1, Pages:="1", PageType:=wdPrintAllPages, _
         ManualDuplexPrint:=False, Collate:=True, Background:=True, PrintToFile:= _
        False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
        PrintZoomPaperHeight:=0
    CommandBars("Control Toolbox").Visible = False        *)


end;
end.

codeunit 89001 "AK Workbook Mgt"
{
    // Facade for saving/loading workbooks and moving files in and out of the add-in.
    // Other extensions should call this codeunit instead of touching table "AK Workbook" directly.

    Access = Public;

    var
        DownloadDialogTxt: Label 'Export workbook';
        ExcelFilterTxt: Label 'Excel Workbook (*.xlsx)|*.xlsx', Locked = true;
        DefaultFileNameTxt: Label 'Workbook', Locked = true;

    procedure OpenWorkbook(WorkbookCode: Code[20])
    var
        Workbook: Record "AK Workbook";
        Editor: Page "AK Workbook Editor";
    begin
        Workbook.Get(WorkbookCode);
        Editor.SetWorkbook(Workbook."Code");
        Editor.Run();
    end;

    procedure GetEditorConfigJson(WorkbookCode: Code[20]): Text
    begin
        exit(GetEditorConfigJson(WorkbookCode, ''));
    end;

    /// <summary>Add-in configuration. AutoRegionJson (optional): a region to show or create at start, sent when the workbook is opened from a BC page.</summary>
    procedure GetEditorConfigJson(WorkbookCode: Code[20]; AutoRegionJson: Text): Text
    var
        Workbook: Record "AK Workbook";
        DataService: Codeunit "AK Workbook Data Service";
        WorkbookLines: Codeunit "AK Workbook Lines";
        Config: JsonObject;
        AutoRegion: JsonObject;
        Result: Text;
    begin
        Config.ReadFrom(DataService.GetConfigJson());
        if Workbook.Get(WorkbookCode) then begin
            Config.Add('workbookCode', Workbook."Code");
            Config.Add('description', Workbook.Description);
        end;
        Config.Add('quickStarts', WorkbookLines.GetQuickStartsJson());
        if AutoRegionJson <> '' then
            if AutoRegion.ReadFrom(AutoRegionJson) then
                Config.Add('autoRegion', AutoRegion);
        Config.WriteTo(Result);
        exit(Result);
    end;

    /// <summary>Opens (and if needed creates) a workbook that shows the given region, e.g. a journal batch or a document's lines.</summary>
    procedure OpenWithRegion(WorkbookCode: Code[20]; Description: Text; Region: JsonObject)
    var
        Workbook: Record "AK Workbook";
        Editor: Page "AK Workbook Editor";
        RegionText: Text;
    begin
        if not Workbook.Get(WorkbookCode) then begin
            Workbook.Init();
            Workbook."Code" := WorkbookCode;
            Workbook.Description := CopyStr(Description, 1, MaxStrLen(Workbook.Description));
            Workbook.Insert(true);
            Commit();
        end;
        Region.WriteTo(RegionText);
        Editor.SetWorkbook(Workbook."Code");
        Editor.SetAutoRegion(RegionText);
        Editor.Run();
    end;

    procedure SetDescription(WorkbookCode: Code[20]; NewDescription: Text)
    var
        Workbook: Record "AK Workbook";
    begin
        Workbook.Get(WorkbookCode);
        Workbook.Description := CopyStr(NewDescription, 1, MaxStrLen(Workbook.Description));
        Workbook.Modify(true);
    end;

    procedure SaveWorkbook(WorkbookCode: Code[20]; Payload: Text; LinkedTables: Text)
    var
        Workbook: Record "AK Workbook";
        OutStr: OutStream;
    begin
        Workbook.Get(WorkbookCode);
        Clear(Workbook."Workbook Data");
        Workbook."Workbook Data".CreateOutStream(OutStr, TextEncoding::UTF8);
        OutStr.WriteText(Payload);
        Workbook."Has Data" := Payload <> '';
        Workbook."Data Size (KB)" := (StrLen(Payload) + 1023) div 1024;
        Workbook."Linked Tables" := CopyStr(LinkedTables, 1, MaxStrLen(Workbook."Linked Tables"));
        Workbook."Last Saved At" := CurrentDateTime();
        Workbook."Last Saved By" := CopyStr(UserId(), 1, MaxStrLen(Workbook."Last Saved By"));
        Workbook.Modify(true);
    end;

    procedure GetWorkbookPayload(WorkbookCode: Code[20]): Text
    var
        Workbook: Record "AK Workbook";
        InStr: InStream;
        Builder: TextBuilder;
        Line: Text;
        LineFeed: Text[1];
        IsFirstLine: Boolean;
    begin
        if not Workbook.Get(WorkbookCode) then
            exit('');
        Workbook.CalcFields("Workbook Data");
        if not Workbook."Workbook Data".HasValue() then
            exit('');

        LineFeed[1] := 10;
        IsFirstLine := true;
        Workbook."Workbook Data".CreateInStream(InStr, TextEncoding::UTF8);
        while not InStr.EOS() do begin
            InStr.ReadText(Line);
            if not IsFirstLine then
                Builder.Append(LineFeed);
            Builder.Append(Line);
            IsFirstLine := false;
        end;
        exit(Builder.ToText());
    end;

    procedure SplitIntoChunks(Payload: Text; ChunkSize: Integer; var Chunks: List of [Text])
    var
        Position: Integer;
        PayloadLength: Integer;
    begin
        Clear(Chunks);
        PayloadLength := StrLen(Payload);
        if PayloadLength = 0 then begin
            Chunks.Add('');
            exit;
        end;
        Position := 1;
        while Position <= PayloadLength do begin
            Chunks.Add(CopyStr(Payload, Position, ChunkSize));
            Position += ChunkSize;
        end;
    end;

    procedure DownloadXlsx(Base64Text: Text; FileName: Text)
    var
        TempBlob: Codeunit "Temp Blob";
        Base64Convert: Codeunit "Base64 Convert";
        OutStr: OutStream;
        InStr: InStream;
        ToFile: Text;
    begin
        ToFile := FileName;
        if ToFile = '' then
            ToFile := DefaultFileNameTxt + '.xlsx';
        TempBlob.CreateOutStream(OutStr);
        Base64Convert.FromBase64(Base64Text, OutStr);
        TempBlob.CreateInStream(InStr);
        DownloadFromStream(InStr, DownloadDialogTxt, '', ExcelFilterTxt, ToFile);
    end;

    procedure GetExportFileName(WorkbookCode: Code[20]; Description: Text): Text
    var
        BaseName: Text;
    begin
        BaseName := Description;
        if BaseName = '' then
            BaseName := WorkbookCode;
        if BaseName = '' then
            BaseName := DefaultFileNameTxt;
        BaseName := DelChr(BaseName, '=', '\/:*?"<>|');
        exit(BaseName + '.xlsx');
    end;

}

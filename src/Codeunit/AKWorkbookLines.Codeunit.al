codeunit 89004 "AK Workbook Lines"
{
    // Stage 2: journal and document lines in a workbook.
    // - Builds the region definitions for General Journal, Item Journal, Sales and Purchase document lines
    //   (fields in page order, filters on template/batch or document).
    // - Opens a workbook with such a region from the BC pages ("Edit in Excel Workbook" actions), and offers
    //   the same definitions as quick-start tiles in the region panel.
    // New rows typed in these regions get their key fields from the region filters (template, batch,
    // document type and no.), the next Line No. and, for journals, the page's SetUpNewLine defaults
    // (see "AK Workbook Write Row").

    Access = Public;

    var
        GenJnlTitleTxt: Label 'General journal';
        ItemJnlTitleTxt: Label 'Item journal';
        SalesLinesTitleTxt: Label 'Sales order lines';
        PurchLinesTitleTxt: Label 'Purchase order lines';
        GenJnlDescTxt: Label 'General Journal %1 / %2', Comment = '%1 = journal template, %2 = journal batch';
        ItemJnlDescTxt: Label 'Item Journal %1 / %2', Comment = '%1 = journal template, %2 = journal batch';
        SalesDescTxt: Label 'Sales %1 %2', Comment = '%1 = document type, %2 = document no.';
        PurchDescTxt: Label 'Purchase %1 %2', Comment = '%1 = document type, %2 = document no.';

    /// <summary>Quick-start tiles for the region panel, as JSON: [{ id, title, region }].</summary>
    procedure GetQuickStartsJson(): JsonArray
    var
        GenJnlTemplate: Record "Gen. Journal Template";
        GenJnlBatch: Record "Gen. Journal Batch";
        ItemJnlTemplate: Record "Item Journal Template";
        ItemJnlBatch: Record "Item Journal Batch";
        SalesHeader: Record "Sales Header";
        PurchHeader: Record "Purchase Header";
        Tiles: JsonArray;
        Template: Code[10];
        Batch: Code[10];
    begin
        GenJnlTemplate.SetRange(Type, GenJnlTemplate.Type::General);
        GenJnlTemplate.SetRange(Recurring, false);
        if GenJnlTemplate.FindFirst() then begin
            Template := GenJnlTemplate.Name;
            GenJnlBatch.SetRange("Journal Template Name", Template);
            if GenJnlBatch.FindFirst() then
                Batch := GenJnlBatch.Name;
        end;
        AddTile(Tiles, 'genjnl', GenJnlTitleTxt, GeneralJournalRegion(Template, Batch));

        Clear(Template);
        Clear(Batch);
        ItemJnlTemplate.SetRange(Type, ItemJnlTemplate.Type::Item);
        ItemJnlTemplate.SetRange(Recurring, false);
        if ItemJnlTemplate.FindFirst() then begin
            Template := ItemJnlTemplate.Name;
            ItemJnlBatch.SetRange("Journal Template Name", Template);
            if ItemJnlBatch.FindFirst() then
                Batch := ItemJnlBatch.Name;
        end;
        AddTile(Tiles, 'itemjnl', ItemJnlTitleTxt, ItemJournalRegion(Template, Batch));

        SalesHeader.SetRange("Document Type", SalesHeader."Document Type"::Order);
        if SalesHeader.FindLast() then;
        AddTile(Tiles, 'sales', SalesLinesTitleTxt, SalesLinesRegion(SalesHeader."Document Type"::Order, SalesHeader."No."));

        PurchHeader.SetRange("Document Type", PurchHeader."Document Type"::Order);
        if PurchHeader.FindLast() then;
        AddTile(Tiles, 'purch', PurchLinesTitleTxt, PurchaseLinesRegion(PurchHeader."Document Type"::Order, PurchHeader."No."));
        exit(Tiles);
    end;

    // ---------------------------------------------------------------- open from BC pages

    procedure OpenGeneralJournal(var GenJnlLine: Record "Gen. Journal Line")
    var
        WorkbookMgt: Codeunit "AK Workbook Mgt";
        TemplateFilter: Text;
        BatchFilter: Text;
        Template: Code[10];
        Batch: Code[10];
    begin
        GenJnlLine.FilterGroup(2);
        TemplateFilter := GenJnlLine.GetFilter("Journal Template Name");
        BatchFilter := GenJnlLine.GetFilter("Journal Batch Name");
        GenJnlLine.FilterGroup(0);
        GetJournalFilters(GenJnlLine."Journal Template Name", GenJnlLine."Journal Batch Name", TemplateFilter, BatchFilter, Template, Batch);
        WorkbookMgt.OpenWithRegion(
            CopyStr('GJ ' + Template + ' ' + Batch, 1, 20), StrSubstNo(GenJnlDescTxt, Template, Batch),
            GeneralJournalRegion(Template, Batch));
    end;

    procedure OpenItemJournal(var ItemJnlLine: Record "Item Journal Line")
    var
        WorkbookMgt: Codeunit "AK Workbook Mgt";
        TemplateFilter: Text;
        BatchFilter: Text;
        Template: Code[10];
        Batch: Code[10];
    begin
        ItemJnlLine.FilterGroup(2);
        TemplateFilter := ItemJnlLine.GetFilter("Journal Template Name");
        BatchFilter := ItemJnlLine.GetFilter("Journal Batch Name");
        ItemJnlLine.FilterGroup(0);
        GetJournalFilters(ItemJnlLine."Journal Template Name", ItemJnlLine."Journal Batch Name", TemplateFilter, BatchFilter, Template, Batch);
        WorkbookMgt.OpenWithRegion(
            CopyStr('IJ ' + Template + ' ' + Batch, 1, 20), StrSubstNo(ItemJnlDescTxt, Template, Batch),
            ItemJournalRegion(Template, Batch));
    end;

    procedure OpenSalesDocument(SalesHeader: Record "Sales Header")
    var
        WorkbookMgt: Codeunit "AK Workbook Mgt";
    begin
        WorkbookMgt.OpenWithRegion(
            CopyStr('S ' + SalesHeader."No.", 1, 20), StrSubstNo(SalesDescTxt, SalesHeader."Document Type", SalesHeader."No."),
            SalesLinesRegion(SalesHeader."Document Type", SalesHeader."No."));
    end;

    procedure OpenPurchaseDocument(PurchHeader: Record "Purchase Header")
    var
        WorkbookMgt: Codeunit "AK Workbook Mgt";
    begin
        WorkbookMgt.OpenWithRegion(
            CopyStr('P ' + PurchHeader."No.", 1, 20), StrSubstNo(PurchDescTxt, PurchHeader."Document Type", PurchHeader."No."),
            PurchaseLinesRegion(PurchHeader."Document Type", PurchHeader."No."));
    end;

    // ---------------------------------------------------------------- region definitions (fields in page order)

    procedure GeneralJournalRegion(Template: Code[10]; Batch: Code[10]): JsonObject
    var
        GenJnlLine: Record "Gen. Journal Line";
        FieldNos: List of [Integer];
        Filters: Dictionary of [Integer, Text];
    begin
        FieldNos.Add(GenJnlLine.FieldNo("Journal Template Name"));
        FieldNos.Add(GenJnlLine.FieldNo("Journal Batch Name"));
        FieldNos.Add(GenJnlLine.FieldNo("Line No."));
        FieldNos.Add(GenJnlLine.FieldNo("Posting Date"));
        FieldNos.Add(GenJnlLine.FieldNo("Document Type"));
        FieldNos.Add(GenJnlLine.FieldNo("Document No."));
        FieldNos.Add(GenJnlLine.FieldNo("Account Type"));
        FieldNos.Add(GenJnlLine.FieldNo("Account No."));
        FieldNos.Add(GenJnlLine.FieldNo(Description));
        FieldNos.Add(GenJnlLine.FieldNo(Amount));
        FieldNos.Add(GenJnlLine.FieldNo("Bal. Account Type"));
        FieldNos.Add(GenJnlLine.FieldNo("Bal. Account No."));
        FieldNos.Add(GenJnlLine.FieldNo("Shortcut Dimension 1 Code"));
        FieldNos.Add(GenJnlLine.FieldNo("Shortcut Dimension 2 Code"));
        Filters.Add(GenJnlLine.FieldNo("Journal Template Name"), Template);
        Filters.Add(GenJnlLine.FieldNo("Journal Batch Name"), Batch);
        exit(BuildRegion(Database::"Gen. Journal Line", FieldNos, Filters));
    end;

    procedure ItemJournalRegion(Template: Code[10]; Batch: Code[10]): JsonObject
    var
        ItemJnlLine: Record "Item Journal Line";
        FieldNos: List of [Integer];
        Filters: Dictionary of [Integer, Text];
    begin
        FieldNos.Add(ItemJnlLine.FieldNo("Journal Template Name"));
        FieldNos.Add(ItemJnlLine.FieldNo("Journal Batch Name"));
        FieldNos.Add(ItemJnlLine.FieldNo("Line No."));
        FieldNos.Add(ItemJnlLine.FieldNo("Posting Date"));
        FieldNos.Add(ItemJnlLine.FieldNo("Entry Type"));
        FieldNos.Add(ItemJnlLine.FieldNo("Document No."));
        FieldNos.Add(ItemJnlLine.FieldNo("Item No."));
        FieldNos.Add(ItemJnlLine.FieldNo(Description));
        FieldNos.Add(ItemJnlLine.FieldNo("Location Code"));
        FieldNos.Add(ItemJnlLine.FieldNo(Quantity));
        FieldNos.Add(ItemJnlLine.FieldNo("Unit Amount"));
        FieldNos.Add(ItemJnlLine.FieldNo(Amount));
        Filters.Add(ItemJnlLine.FieldNo("Journal Template Name"), Template);
        Filters.Add(ItemJnlLine.FieldNo("Journal Batch Name"), Batch);
        exit(BuildRegion(Database::"Item Journal Line", FieldNos, Filters));
    end;

    procedure SalesLinesRegion(DocumentType: Enum "Sales Document Type"; DocumentNo: Code[20]): JsonObject
    var
        SalesLine: Record "Sales Line";
        FieldNos: List of [Integer];
        Filters: Dictionary of [Integer, Text];
    begin
        FieldNos.Add(SalesLine.FieldNo("Document Type"));
        FieldNos.Add(SalesLine.FieldNo("Document No."));
        FieldNos.Add(SalesLine.FieldNo("Line No."));
        FieldNos.Add(SalesLine.FieldNo(Type));
        FieldNos.Add(SalesLine.FieldNo("No."));
        FieldNos.Add(SalesLine.FieldNo(Description));
        FieldNos.Add(SalesLine.FieldNo("Location Code"));
        FieldNos.Add(SalesLine.FieldNo(Quantity));
        FieldNos.Add(SalesLine.FieldNo("Unit of Measure Code"));
        FieldNos.Add(SalesLine.FieldNo("Unit Price"));
        FieldNos.Add(SalesLine.FieldNo("Line Discount %"));
        FieldNos.Add(SalesLine.FieldNo("Line Amount"));
        Filters.Add(SalesLine.FieldNo("Document Type"), Format(DocumentType));
        Filters.Add(SalesLine.FieldNo("Document No."), DocumentNo);
        exit(BuildRegion(Database::"Sales Line", FieldNos, Filters));
    end;

    procedure PurchaseLinesRegion(DocumentType: Enum "Purchase Document Type"; DocumentNo: Code[20]): JsonObject
    var
        PurchLine: Record "Purchase Line";
        FieldNos: List of [Integer];
        Filters: Dictionary of [Integer, Text];
    begin
        FieldNos.Add(PurchLine.FieldNo("Document Type"));
        FieldNos.Add(PurchLine.FieldNo("Document No."));
        FieldNos.Add(PurchLine.FieldNo("Line No."));
        FieldNos.Add(PurchLine.FieldNo(Type));
        FieldNos.Add(PurchLine.FieldNo("No."));
        FieldNos.Add(PurchLine.FieldNo(Description));
        FieldNos.Add(PurchLine.FieldNo("Location Code"));
        FieldNos.Add(PurchLine.FieldNo(Quantity));
        FieldNos.Add(PurchLine.FieldNo("Unit of Measure Code"));
        FieldNos.Add(PurchLine.FieldNo("Direct Unit Cost"));
        FieldNos.Add(PurchLine.FieldNo("Line Discount %"));
        FieldNos.Add(PurchLine.FieldNo("Line Amount"));
        Filters.Add(PurchLine.FieldNo("Document Type"), Format(DocumentType));
        Filters.Add(PurchLine.FieldNo("Document No."), DocumentNo);
        exit(BuildRegion(Database::"Purchase Line", FieldNos, Filters));
    end;

    /// <summary>Region JSON as the add-in stores it: tableNo, tableCaption, fields[{no, caption, type, pk}], filters[{fieldNo, filter}], editable.</summary>
    procedure BuildRegion(TableNo: Integer; FieldNos: List of [Integer]; Filters: Dictionary of [Integer, Text]): JsonObject
    var
        WriteService: Codeunit "AK Workbook Write Service";
        RecRef: RecordRef;
        PrimaryKey: KeyRef;
        PkFields: List of [Integer];
        Region: JsonObject;
        FieldsJson: JsonArray;
        FiltersJson: JsonArray;
        FieldJson: JsonObject;
        FilterJson: JsonObject;
        FieldNo: Integer;
        Index: Integer;
    begin
        RecRef.Open(TableNo);
        PrimaryKey := RecRef.KeyIndex(1);
        for Index := 1 to PrimaryKey.FieldCount() do
            PkFields.Add(PrimaryKey.FieldIndex(Index).Number());
        foreach FieldNo in FieldNos do
            if RecRef.FieldExist(FieldNo) then begin
                Clear(FieldJson);
                FieldJson.Add('no', FieldNo);
                FieldJson.Add('caption', RecRef.Field(FieldNo).Caption());
                FieldJson.Add('type', Format(RecRef.Field(FieldNo).Type()));
                FieldJson.Add('pk', PkFields.Contains(FieldNo));
                FieldsJson.Add(FieldJson);
            end;
        foreach FieldNo in Filters.Keys() do
            if Filters.Get(FieldNo) <> '' then begin
                Clear(FilterJson);
                FilterJson.Add('fieldNo', FieldNo);
                FilterJson.Add('filter', Filters.Get(FieldNo));
                FiltersJson.Add(FilterJson);
            end;
        Region.Add('tableNo', TableNo);
        Region.Add('tableCaption', RecRef.Caption());
        Region.Add('fields', FieldsJson);
        Region.Add('filters', FiltersJson);
        Region.Add('editable', WriteService.CanModify(TableNo));
        Region.Add('editReason', WriteService.GetEditBlockReason(TableNo));
        RecRef.Close();
        exit(Region);
    end;

    local procedure AddTile(var Tiles: JsonArray; Id: Text; Title: Text; Region: JsonObject)
    var
        Tile: JsonObject;
    begin
        Tile.Add('id', Id);
        Tile.Add('title', Title);
        Tile.Add('region', Region);
        Tiles.Add(Tile);
    end;

    /// <summary>The journal pages filter on template and batch (filter group 2); fall back to the current line.</summary>
    local procedure GetJournalFilters(RecTemplate: Code[10]; RecBatch: Code[10]; TemplateFilter: Text; BatchFilter: Text; var Template: Code[10]; var Batch: Code[10])
    begin
        Template := RecTemplate;
        Batch := RecBatch;
        if TemplateFilter <> '' then
            Template := CopyStr(DelChr(TemplateFilter, '=', ''''), 1, MaxStrLen(Template));
        if BatchFilter <> '' then
            Batch := CopyStr(DelChr(BatchFilter, '=', ''''), 1, MaxStrLen(Batch));
    end;
}

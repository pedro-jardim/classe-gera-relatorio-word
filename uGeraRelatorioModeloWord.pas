unit uGeraRelatorioModeloWord;

interface

Uses
  Classes, Word2000, ComObj, OleServer, SysUtils, StdCtrls,
  Windows, Variants, Funcoes, DB, ExtCtrls, Clipbrd, Printers, DBCtrls;
  
type
  TDataSetItem = class(TObject)
    private
      FDataSetItem   : TDataSet;
      FTableIndex    : Integer;
      FTableIndexPai : Integer;
      FCampoAgrupado : string;
    public
      property DataSetItem   : TDataSet read FDataSetItem   write FDataSetItem;
      property Tableindex    : Integer  read FTableindex    write FTableIndex;
      property TableIndexPai : Integer  read FTableIndexPai write FTableIndexPai;
      property CampoAgrupado : String   read FCampoAgrupado write FCampoAgrupado;
    end;


type
  TListaDataSetItens = class(TObject)
    private
      voLista : TList;
      function GetDataSetItem( const pIndex : Integer ) : TDataSetItem;
      procedure SetDataSetItem( const pIndex : Integer; const Value : TDataSetItem );
    public
      property Itens[ const pIndex : Integer ] : TDataSetItem read GetDataSetItem write SetDataSetItem;
      function Add : TDataSetItem;
      function Count : Integer;
      procedure Clear;
      procedure Delete( const pIndex : Integer );
      constructor Create;
      destructor Free;
    end;

type
  TRelatorioWord = Class
    private
    //propriedades
      FDiretorioDestino   : String;
      FDiretorioModelo    : String;
      FWordApplication    : Variant;
      FNomeRelatorio      :  String;
      FVisivel            : Boolean;
      FDataSetCapa        : TDataSet;
      FDataSetItens       : TDataSet;
      FDataSetPadrao      : TDataSet;
      FRodape             : Boolean;
      FCabecalho          : Boolean;
      FCampoQuantidade    : String;
      FNomeImpressora     : String;
      voListaDataSetItens : TListaDataSetItens;
      FImprimir           : Boolean;
      FMemo               : TDBMemo;
      FDataSourceCapa     : TDataSource;
      FImage              : TImage;



      //--metodos
      function TrataDiretorios(var pMsg: String; const pCopia: Boolean ): Boolean;
      function GeraDadosCapa(var Msg : String; Doc : Variant): Boolean;
      function GeraDadosTabela(var Msg : String; Doc : Variant): Boolean;
      function GeraDadosEtiqueta(var pMsg : String; pQtde: Integer): Boolean;
      procedure PopulaTabela(var pIndex : Integer; Doc : Variant);
      procedure PopulaTabelaAgrupada(var pIndex : Integer; Doc : Variant);
      procedure InsereImagem;

    public
      //--propriedades
      property DiretorioModelo  : String             read FDiretorioModelo      write FDiretorioModelo;
      property DiretorioDestino : String             read FDiretorioDestino     write FDiretorioDestino;
      property Visivel          : Boolean            read FVisivel              write FVisivel;
      property DataSetPadrao    : TDataSet           read FDataSetPadrao        write FDataSetPadrao;
      property DataSetCapa      : TDataSet           read FDataSetCapa          write FDataSetCapa;
      property Cabecalho        : Boolean            read FCabecalho            write FCabecalho;
      property Rodape           : Boolean            read FRodape               write FRodape;
      property CampoQuantidade  : String             read FCampoQuantidade      write FCampoQuantidade;
      property DataSetItens     : TListaDataSetItens read voListaDataSetItens   write voListaDataSetItens;
      property NomeImpressora   : String             read FNomeImpressora       write FNomeImpressora;
      property Imprimir         : Boolean            read FImprimir             write FImprimir;
      property Memo             : TDBMemo            read FMemo                 write FMemo;
      property DataSourceCapa   : TDataSource        read FDataSourceCapa       write FDataSourceCapa;
      property Image            : TImage             read FImage                write FImage;

      //--métodos
      function GeraRelatorio(var Msg: String): Boolean;
      function GeraEtiqueta(var pMsg: String; pQtde: Integer): Boolean;

      //--métodos construtores e destrutores
      constructor Create;
      destructor  Free;
    end;


implementation

{ TRelatorioWord }

constructor TRelatorioWord.Create;
begin
  FWordApplication := CreateOleObject('Word.Application');
  FVisivel            := False;
  FCabecalho          := false;
  FRodape             := false;
  voListaDataSetItens := TListaDataSetItens.Create;
end;

destructor TRelatorioWord.Free;
begin
  voListaDataSetItens.Free;

  FWordApplication.Quit;
  FWordApplication := unassigned;
end;

function TRelatorioWord.GeraDadosEtiqueta(var pMsg: String; pQtde: Integer): Boolean;
var
  lDoc   : Variant;
  lCount,
  lQtde,
  lQtdeTotal : Integer;
  lValor,
  lCampo : String;
  lRange : Variant;
begin
  Result := True;



  try

    with DataSetPadrao do
    begin
      First;
      if not Eof then
      begin
        while not Eof do
        begin
          if pQtde > 0 then
            lQtdeTotal := pQtde
          else
            lQtdeTotal := DataSetPadrao.FieldByName(CampoQuantidade).AsInteger;

          for lQtde := 1 to lQtdeTotal do
          begin
            lDoc  := FWordApplication.Documents.Open(DiretorioModelo);
            for lcount := 0 to FieldCount - 1 do
            begin
              lValor := '';
              lcampo := FieldList.Fields[lcount].FieldName;
              if FieldList.Fields[lcount].DataType = ftFloat then
                lvalor := floattostrf(FieldList.Fields[lcount].asfloat,ffnumber,10,3) //Modificado para o fagner
              else if UpperCase(Copy(FieldList.Fields[lcount].FieldName,1,3)) = 'IMG' then
                begin

                  lRange := lDoc.Content.Find.Execute(FindText := '['+lcampo+']', ReplaceWith := '', Replace := wdReplaceAll);

                  if FileExists(FieldList.Fields[lcount].AsString) then
                  begin
                    lDoc.InlineShapes.AddPicture(
                              FileName := FieldList.Fields[lcount].AsString,
                              Range    := lRange,
                      SaveWithDocument := True )
                  end;

                end
              else
                lvalor := FieldList.Fields[lcount].AsString;
              if Length(lvalor) > 255 then
              begin
                Clipboard .AsText := lvalor;

                lDoc.Content.Find.Execute(FindText := '['+lcampo+']', ReplaceWith := '^c', Replace := wdReplaceAll);

                Clipboard.Clear;
              end
              else
              begin
                lDoc.Content.Find.Execute(FindText := '['+lcampo+']', ReplaceWith := lvalor, Replace := wdReplaceAll);

              end;
            end;

            //Define a impressora que o Word irá imprimir, caso não informado usa a padrão
            if NomeImpressora <> '' then
              FWordApplication.ActivePrinter := NomeImpressora;

            //Silencia os alertas de impressão do word
            FWordApplication.DisplayAlerts := false;

            lDoc.PrintOut;
            lDoc.Close(False);

          end;

          Next;
        end;
      end
      else
      begin
        pMsg    := 'Não existem dados no DataSet Padrão';
        Result := False;
      end;

    end;

  except
    on E : Exception do
    begin
      pMsg    := 'Erro na geração dos dados da etiqueta';
      Result := False;
    end;
  end;

end;

function TRelatorioWord.GeraDadosCapa(var Msg: String;
                                          Doc: Variant): Boolean;
var
  count        : integer;
  campo, valor : string;
  lRange : Variant;
begin
  result := true;
  with DataSetCapa do
    begin
      if not eof then
        begin
          First;

          //habilita o cabeçalho
          if FCabecalho then
            Doc.ActiveWindow.ActivePane.View.SeekView := 9;
          //habilita o rodapé
          if FRodape then
            Doc.activewindow.Activepane.view.seekview := 10;


          for count := 0 to FieldCount - 1 do
            begin
              campo := FieldList.Fields[count].FieldName;
              if FieldList.Fields[count].DataType = ftFloat then
                valor := floattostrf(FieldList.Fields[count].asfloat,ffnumber,10,2)
              else if UpperCase(Copy(FieldList.Fields[count].FieldName,1,3)) = 'IMG' then
              begin
                if FileExists(FieldList.Fields[count].AsString) then
                begin
                  //Doc.InlineShapes.AddPicture(
                  //          FileName := FieldList.Fields[count].AsString,
                  //          Range    := lRange,
                  //  SaveWithDocument := True )
                  FImage.Picture := nil;
                  FImage.Picture.LoadFromFile(FieldList.Fields[count].AsString);
                  Clipboard.Assign(FImage.Picture.Graphic);

                  Doc.Content.Find.Execute(FindText := '['+campo+']', ReplaceWith := '^c', Replace := wdReplaceAll);

                  //Preenche Cabeçalho
                  if Cabecalho then
                    Doc.Sections.Item(1).Headers.Item(wdHeaderFooterPrimary).Range.Find
                                .Execute(FindText := '['+campo+']', ReplaceWith := '^c', Replace := wdReplaceAll);
                  //Preenche Rodapé
                  if rodape then
                    Doc.Sections.Item(1).Footers.Item(wdHeaderFooterPrimary).Range.Find
                                .Execute(FindText := '['+campo+']', ReplaceWith := '^c', Replace := wdReplaceAll);

                  Clipboard.Clear;
                end;

              end

              else if FieldList.Fields[count].DataType = ftMemo then
              begin
                if Assigned(FDataSourceCapa) and Assigned(FMemo) then
                begin
                  //Vincula a procedure com um dataSource
                  FDataSourceCapa.DataSet := FDataSetCapa;
                  //Faz com que coloque o campo em tempo de execução dentro do MEMO
                  FMemo.DataSource := FDataSourceCapa;
                  FMemo.DataField  := FDataSetCapa.Fieldbyname(campo).FieldName;
                  valor := FMemo.Lines.Text;
                end
              end
              else
                valor := FieldList.Fields[count].AsString;
              if Length(valor) > 255 then
              begin
                Clipboard.AsText := valor;

                Doc.Content.Find.Execute(FindText := '['+campo+']', ReplaceWith := '^c', Replace := wdReplaceAll);

                //Preenche Cabeçalho
                if Cabecalho then
                  Doc.Sections.Item(1).Headers.Item(wdHeaderFooterPrimary).Range.Find
                              .Execute(FindText := '['+campo+']', ReplaceWith := '^c', Replace := wdReplaceAll);
                //Preenche Rodapé
                if rodape then
                  Doc.Sections.Item(1).Footers.Item(wdHeaderFooterPrimary).Range.Find
                              .Execute(FindText := '['+campo+']', ReplaceWith := '^c', Replace := wdReplaceAll);

                Clipboard.Clear;
              end
              else
              begin
                Doc.Content.Find.Execute(FindText := '['+campo+']', ReplaceWith := valor, Replace := wdReplaceAll);
                //Preenche Cabeçalho
                if Cabecalho then
                  Doc.Sections.Item(1).Headers.Item(wdHeaderFooterPrimary).Range.Find
                              .Execute(FindText := '['+campo+']', ReplaceWith := valor, Replace := wdReplaceAll);
                //Preenche Rodapé
                if rodape then
                  Doc.Sections.Item(1).Footers.Item(wdHeaderFooterPrimary).Range.Find
                              .Execute(FindText := '['+campo+']', ReplaceWith := valor, Replace := wdReplaceAll);
              end;
            end;
            //Define a impressora que o Word irá imprimir, caso não informado usa a padrão
            if NomeImpressora <> '' then
              FWordApplication.ActivePrinter := NomeImpressora;
        end
      else
        begin
          Msg    := 'Não existem dados no dataset da capa';
          Result := False;
      end;
    end;


end;

function TRelatorioWord.GeraDadosTabela(var Msg: String;
                                            Doc: Variant): Boolean;
var
  i : Integer;
begin
  result := True;
  
  try

    for i := 0 to DataSetItens.Count - 1 do
    begin
      //--> Verifica o formato de popular a tabela
      if DataSetItens.Itens[i].CampoAgrupado = '' then
        PopulaTabela(i, Doc)
      else
        PopulaTabelaAgrupada(i, Doc)
    end;

  except
    on E : Exception do
    begin
      Result := False;
      Msg := 'Problema ao gravar os itens.'+#13+E.Message;
    end
  end;

end;

function TRelatorioWord.GeraEtiqueta(var pMsg: String; pQtde: Integer ): Boolean;
var
  lDoc : Variant;
begin
  pMsg := EmptyStr;

  FWordApplication.Visible := Visivel;

  Result := TrataDiretorios(pMsg, False);

  if Result then
    Result := GeraDadosEtiqueta(pMsg, pQtde);

end;

function TRelatorioWord.GeraRelatorio(var Msg: String): Boolean;
var
  Doc : Variant;
begin
  msg := EmptyStr;

  FwordApplication.Visible := Visivel;

  Result := TrataDiretorios(msg, True);

  if Result then
  begin
    Doc := FWordApplication.Documents.Open(DiretorioDestino);

    result := GeraDadosCapa(msg,doc);

    if Result then
      Result := GeraDadosTabela(msg, doc);

    if Imprimir then
      Doc.Printout(False);

    if Result then
      Doc.Save;

  end;

end;

function TRelatorioWord.TrataDiretorios(var pMsg: String; const pCopia: Boolean ): Boolean;
begin
  result := True;
  try

    if FileExists(FDiretorioModelo) then
    begin

      if pCopia then
      begin
        if DirectoryExists(ExtractFilePath(DiretorioDestino)) then
        begin
          CopyFile(pchar( trim(DiretorioModelo) ),pchar( DiretorioDestino ),false);
          Result := true;
        end
        else
        begin
          pMsg   := 'Diretório de Destino não encontrado.';
          Result := False;
        end;
      end;

    end
    else
    begin
      pMsg := 'Modelo não encontrado no diretório';
      Result := False;
    end;

  except
    on E : Exception do
    begin
      pMsg := 'Problema ao tratar os diretórios.' + #13 + E.Message;
      Result := False;
    end;
  end;
end;


procedure TRelatorioWord.InsereImagem;
begin

end;


procedure TRelatorioWord.PopulaTabela(var pIndex : Integer; Doc : Variant);
var
  count, countCelula,
  countLinha             : Integer;
  table, tablepai, linha : Variant;
  nomeCol, valor         : String;
  FontDaTable            : string;
  SizeDaTable            : Integer;
begin

  with DataSetItens.Itens[pIndex].DataSetItem do
  begin
    if not eof then
    begin
      First;

      //-->>Verifica se tem tabela pai
      if DataSetItens.Itens[pIndex].TableIndexPai > 0 then
      begin
        TablePai := Doc.Tables.Item(DataSetItens.Itens[pIndex].TableindexPai);

        //-->> Selecionando a Tabela Filha através da Pai
        Table := TablePai.Tables.Item(DataSetItens.Itens[pIndex].Tableindex);
      end
      else
      begin
        Table := Doc.Tables.Item(DataSetItens.Itens[pIndex].Tableindex);
      end;

      //-->>pega a cópia da linha
      linha       := Table.Rows.Item(2);
      FontDaTable := linha.range.font.Name;
      SizeDaTable := linha.range.font.size;

      //-->>Contador de linhas que serão criadas
      countLinha := 0;

      while not Eof do
      begin
        //-->> cria nova linha
        Inc(countLinha);
        Next;
      end;

      for count := 1 to countLinha - 1 do
        Table.Rows.Add(linha);

      First;

      CountLinha := 2;

      //-->>Loop dos registros adicionando valores as linha
      while not eof do
      begin
        //-->>Inicializa o contador de celula
        countCelula := 1;

        //-->>Loop dos campos do data set
        for count := 0 to FieldCount -1 do
        begin
          if FieldList.Fields[count].DataType = ftFloat then
          begin
            Table.Cell(countLinha, countCelula).range.text := floattostrf(FieldList.Fields[count].asfloat,ffnumber,10,2);
            Table.Cell(countLinha, countCelula).range.font.name := FontDaTable;
            Table.Cell(countLinha, countCelula).range.font.size := SizeDaTable;
          end
          else if UpperCase(Copy(FieldList.Fields[count].FieldName,1,3)) = 'IMG' then
            begin
              if FileExists(FieldList.Fields[count].AsString) then
              begin
                Doc.InlineShapes.AddPicture(
                          FileName := FieldList.Fields[count].AsString,
                          Range    := Table.Cell(countLinha,countCelula).range,
                  SaveWithDocument := True )
              end
              else
                Table.Cell(countLinha,countCelula).range.text := '';
            end
          else
          begin
            Table.Cell(countLinha, countCelula).range.text := FieldList.Fields[count].AsString;
            Table.Cell(countLinha, countCelula).range.font.name := FontDaTable;
            Table.Cell(countLinha, countCelula).range.font.size := SizeDaTable;
          end;
          inc(countCelula);

        end;

        Inc(countlinha);

        Next;
      end;
    end;
  end;
end;

procedure TRelatorioWord.PopulaTabelaAgrupada(var pIndex: Integer; Doc: Variant);
var
  count, countCelula,
  countLinha, countReg,
  countGrupo           : Integer;
  linha                : Variant;
  nomeAgrupamento      : String;
  TableModelo, TablePai: Variant;
  ListaAgrupamento     : TStringList;
  iLinhaGrupo, iLinhaTitulo, iLinha : Integer;
  lRange : OleVariant;

begin

  ListaAgrupamento := TStringList.Create;
  try
    with DataSetItens.Itens[pIndex].DataSetItem do
    begin
      if not eof then
      begin

        //--> Relaciona todas os grupos que serão criados ----------------------------------------------------------
        First;

        ListaAgrupamento.Add(FieldByName(DataSetItens.Itens[pIndex].CampoAgrupado).AsString);
        nomeAgrupamento := FieldByName(DataSetItens.Itens[pIndex].CampoAgrupado).AsString;

        while not Eof do
        begin
          if nomeAgrupamento <> FieldByName(DataSetItens.Itens[pIndex].CampoAgrupado).AsString then
          begin
            ListaAgrupamento.Add(FieldByName(DataSetItens.Itens[pIndex].CampoAgrupado).AsString);
            nomeAgrupamento := FieldByName(DataSetItens.Itens[pIndex].CampoAgrupado).AsString;
          end;
          Next;
        end;

        //----------------------------------------------------------------------------------------------------------

        //--> Define a cópia da tabela modelo
        //-->>Verifica se tem tabela pai
        if DataSetItens.Itens[pIndex].TableIndexPai > 0 then
        begin
          TablePai := Doc.Tables.Item(DataSetItens.Itens[pIndex].TableindexPai);

          //-->> Selecionando a Tabela Filha através da Pai
          TableModelo := TablePai.Tables.Item(DataSetItens.Itens[pIndex].Tableindex);
        end
        else
        begin
          TableModelo := Doc.Tables.Item(DataSetItens.Itens[pIndex].Tableindex);
        end;

          //TableModelo := Doc.Tables.Item(DataSetItens.Itens[pIndex].Tableindex);

        lRange := TableModelo.Range;

        lRange.Copy;

        countLinha := 3;

        //--> Loop dos Grupos Encontrados
        for countGrupo := 0 to ListaAgrupamento.Count - 1 do
        begin

          //--> Filtro dos registros pelo campo de agrupamento
          Filtered := False;
          Filter   := DataSetItens.Itens[pIndex].CampoAgrupado+'='+QuotedStr(ListaAgrupamento.Strings[countGrupo]);
          Filtered := True;

          iLinhaGrupo  := 1;
          iLinhaTitulo := 2;
          iLinha       := 3;


          if countGrupo = 0 then
            TableModelo.Cell(1, 1).range.text := ListaAgrupamento.Strings[countGrupo]
          else
            TableModelo.Cell(countLinha-2, 1).range.text := ListaAgrupamento.Strings[countGrupo];

          First;
          while not Eof do
          begin

            //-->>Inicializa o contador de celula
            countCelula := 1;

            //-->>Loop dos campos do data set
            for count := 1 to DataSetItens.Itens[pIndex].DataSetItem.FieldCount -1 do
            begin

              if DataSetItens.Itens[pIndex].DataSetItem.FieldList.Fields[count].DataType = ftFloat then
                TableModelo.Cell(countLinha, countCelula).range.text := floattostrf(DataSetItens.Itens[pIndex].DataSetItem.FieldList.Fields[count].asfloat,ffnumber,10,2)
              else if UpperCase(Copy(DataSetItens.Itens[pIndex].DataSetItem.FieldList.Fields[count].FieldName,1,3)) = 'IMG' then
                begin
                  if FileExists(DataSetItens.Itens[pIndex].DataSetItem.FieldList.Fields[count].AsString) then
                  begin
                    Doc.InlineShapes.AddPicture(
                              FileName := DataSetItens.Itens[pIndex].DataSetItem.FieldList.Fields[count].AsString,
                              Range    := TableModelo.Cell(countLinha,countCelula).range,
                      SaveWithDocument := True )
                  end
                  else
                    TableModelo.Cell(countLinha,countCelula).range.text := '';
                end
              else
                TableModelo.Cell(countLinha, countCelula).range.text := DataSetItens.Itens[pIndex].DataSetItem.FieldList.Fields[count].AsString;

              inc(countCelula);

            end;

            tableModelo.Rows.Add();
            Inc(countlinha);
            Next;

            //--> Se chegou no final, fazer o tratamento para colar no novo range da tabela
            if Eof then
            begin

              iLinhaGrupo := countLinha;

              lRange.SetRange(TableModelo.Rows.Item(countLinha).Range.Start,TableModelo.Rows.Item(countLinha).Range.End);

              //Cola a tabela modelo apenas quando 
              if ListaAgrupamento.Count -1 > countGrupo then
              begin
                lRange.Paste;

                //-->Inclementa a variavel de linha, pois ao colar a tabela modelo, são adicionados mais 3 novas linhas
                //--> Inc(2) pois no final do loop é adicionado um a variavel
                Inc(countLinha,2);

                //--> No fim do loop é adicionado uma nova linha na tabela, que serve de Range para a nova tabela colada
                //--> Dessa forma deve se deletar a linha criada com a nova com o novo contador de linha
                tableModelo.Rows.Item(countLinha).Delete;
              end
              else
                tableModelo.Rows.Item(countLinha).Delete;

            end;

          end;

        end;

      end;

    end;
  finally
    ListaAgrupamento.Free;
  end;

end;

{ TListaDataSetItens }

function TListaDataSetItens.Add: TDataSetItem;
begin

  Result := TDataSetItem.Create;
  Result.CampoAgrupado := '';

  voLista.Add( Result );
end;

procedure TListaDataSetItens.Clear;
begin
  while Count > 0 do
    Delete( 0 );

end;

function TListaDataSetItens.Count: Integer;
begin
  Result := voLista.Count;
end;

constructor TListaDataSetItens.Create;
begin
  voLista := TList.Create;
end;

procedure TListaDataSetItens.Delete(const pIndex: Integer);
var
  loSingle : TDataSetItem;
begin
  loSingle := voLista.Items[ pIndex ];
  loSingle.Free;
  voLista.Delete( pIndex );
end;

destructor TListaDataSetItens.Free;
var
  i : Integer;
begin
  //--> Limpa os datasets da orgigem
  for i := 0 to Count - 1 do
  begin
    Itens[i].DataSetItem.Filter   := '';
    Itens[i].DataSetItem.Filtered := False;
  end;

  Clear;
  voLista.Free;
end;

function TListaDataSetItens.GetDataSetItem(const pIndex: Integer): TDataSetItem;
begin
  Result := voLista.Items[ pIndex ];
end;

procedure TListaDataSetItens.SetDataSetItem(const pIndex: Integer; const Value: TDataSetItem);
begin
  voLista.Items[ pIndex ] := Value;
end;

end.

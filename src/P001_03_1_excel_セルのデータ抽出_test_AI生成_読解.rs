use calamine::{Data, DataType, Range, Reader, Xlsx, open_workbook};

//検索ワードの構造体
struct TargetQuery {
    row_name: String, 
    col_name: String, 
}


struct ExcelScraper {
    //Range<DataType>だとエラーだった。理由はDataTypeがトレイトだからだそうだ
    //Dataはenum構造になっている。enumは<>でくくれるんだ？へえー
    sheet: Range<Data>,
}

impl ExcelScraper {
    //&strのほうがStringより軽いらしいぞ
    //Result<Self, Box<dyn std::error::Error>>のResultは、エラーの分岐も含めた返り値を返す
    //問題なければSelf, エラーならBox<dyn std::error::Error>を返す
    //dynは「～ならなんでも」という意味らしい
    fn new(path: &str) -> Result<Self, Box<dyn std::error::Error>> {
        //「?」はopen_workbook(path)を実行してOkなら続行、駄目ならそのエラーをreturnして終了する分岐機能だそうだ
        let mut workbook: Xlsx<_> = open_workbook(path)?;
        //ok_or()も?に似てるね。駄目なら文章を出力するのだろう
        //ok_or()はエラー値を作る関数らしい。エラー値と共にメッセージが出るのかな？
        //?やok_or()はResult型にしか使えない？
        let range = workbook.worksheet_range_at(0)
            .ok_or("シートがありません")??;
        
        //このfnの返り値がResult型なので、Ok()で返しているのだろう
        //なんでResult型で返す必要があるんだろう？
        Ok(ExcelScraper { sheet: range })
    }

    //Option?誰だお前は
    fn find_column_index(&self, header_name: &str) -> Option<usize> {
        //ここ、なんで10行なんだっけ？
        for row_idx in 0..10 {
            //Some()はOk()に似ている。
            //Ok()がデータかエラーかなら、Some()はデータかNULLか
            //あー、だからifで結んでいるのか。let Someは分岐するからか
            //nth()？誰だお前
            if let Some(row) = self.sheet.rows().nth(row_idx) {
                for (col_idx, cell) in row.iter().enumerate() {
                    if let Some(text) = cell.as_string() {
                        if text.trim() == header_name {
                            //ここ、Someで日和る必要ある？
                            //返り値がOption<usize>だからか？
                            return Some(col_idx);
                        }
                    }
                }
            }
        }
        //SomeのNULLのときをここで回収
        None
    }

    fn find_row_index(&self, item_name: &str) -> Option<usize> {
        //ここはenumerate()で数えてるな
        for (row_idx, row) in self.sheet.rows().enumerate() {
            //なんで3列なんだっけ？
            for col_idx in 0..3 {
                //なんで行と列の番号の入手するfnの書き方が違うんだ？
                if let Some(cell) = row.get(col_idx) {
                    if let Some(text) = cell.as_string() {
                        if text.trim() == item_name {
                            return Some(row_idx);
                        }
                    }
                }
            }
        }
        None
    }

    fn get_value_at(&self, row_idx: usize, col_idx: usize) -> Option<&Data> {
        //get()は指定したセルのデータを入手するということかな？
        self.sheet.get((row_idx, col_idx))
    }
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = "202606929中小企業実態基本調査_個人企業.xlsx";
    
    let scraper = ExcelScraper::new(path)?;
    println!("シートの読み込みに成功。総行数: {}, 総列数: {}", scraper.sheet.height(), scraper.sheet.width());

    let query = TargetQuery {
        //to_string()は新しくStringを生成する。メモリを使うらしいぞ
        row_name: "売上高".to_string(),
        col_name: "情報サービス業".to_string(),
    };

    //検索をmatchでやるのかっこいいな
    let col_idx = match scraper.find_column_index(&query.col_name) {
        Some(idx) => idx,
        None => {
            println!("❌ 「{}」という業界が見つかりませんでした。", query.col_name);
            return Ok(());
        }
    };

    let row_idx = match scraper.find_row_index(&query.row_name) {
        Some(idx) => idx,
        None => {
            println!("❌ 「{}」という項目が見つかりませんでした。", query.row_name);
            return Ok(());
        }
    };

    if let Some(value) = scraper.get_value_at(row_idx, col_idx) {
        println!("🎯 探索完了！");
        println!("【行】{} (インデックス: {})", query.row_name, row_idx);
        println!("【列】{} (インデックス: {})", query.col_name, col_idx);
        println!("【交差点の値】: {:?}", value);
    } else {
        println!("❌ 交差点 ({}, {}) にデータが存在しませんでした。", row_idx, col_idx);
    }

    Ok(())
}
//P001_03_1_excel_セルのデータ抽出_test_AI生成_読解
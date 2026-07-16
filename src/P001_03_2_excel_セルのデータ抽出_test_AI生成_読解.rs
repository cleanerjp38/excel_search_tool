use calamine::{Reader, Xlsx, open_workbook, Data, DataType};

// ==========================================
// 【Step 3】名詞（構造体）の定義
// ==========================================

/// 探したい「行」と「列」の条件をまとめる構造体
struct TargetQuery {
    row_name: String,
    col_name: String,
}

/// Excelの生々しい型（Rangeやライフタイム）を隠蔽する関所
struct ExcelScraper {
    // 修正箇所：DataType（規格）ではなく、Data（実体）を保持する
    //このコードではRange<Data>でなくVecでやっている：俺
    rows: Vec<Vec<Data>>,
}

// ==========================================
// 【Step 3】物語の記述（implによる振る舞い）
// ==========================================

impl ExcelScraper {
    /// 関所の構築：Excelを開き、すべてのデータを標準のVecに写し取ってから閉じる
    fn new(path: &str) -> Result<Self, Box<dyn std::error::Error>> {
        let mut workbook: Xlsx<_> = open_workbook(path)?;
        let range = workbook.worksheet_range_at(0).ok_or("シートがありません")??;
        
        // RangeからVec<Vec<Data>>へ変換
        //ここでひと手間しているから、速度やメモリの負担分が損？：俺
        let mut rows = Vec::new();
        for row in range.rows() {
            rows.push(row.to_vec());
        }
        
        Ok(ExcelScraper { rows })
    }

    /// 列（ターゲット業界）のインデックス特定
    fn find_col(&self, target: &str) -> Option<usize> {
        //rowsがRangeでなくてVecだから、ここも書き方がちがう：俺
        //min()はx<=10と制限する機能なのかな？：俺
        let limit = self.rows.len().min(10); // 上から10行目まで探す
        for row_idx in 0..limit {
            //Rangeだとrowで済むのが、Vecだとrows[row_idx]と書く必要がある：俺
            for (col_idx, cell) in self.rows[row_idx].iter().enumerate() {
                // as_string() は DataType トレイトの機能
                if let Some(text) = cell.as_string() {
                    if text.trim() == target {
                        return Some(col_idx);
                    }
                }
            }
        }
        None
    }

    /// 行（ターゲット項目）のインデックス特定
    fn find_row(&self, target: &str) -> Option<usize> {
        for (row_idx, row) in self.rows.iter().enumerate() {
            let limit = row.len().min(3); // 左から3列目まで探す
            for col_idx in 0..limit {
                if let Some(text) = row[col_idx].as_string() {
                    if text.trim() == target {
                        return Some(row_idx);
                    }
                }
            }
        }
        None
    }

    /// 交差点の値を狙い撃つ
    // 修正箇所：戻り値も &DataType ではなく &Data
    fn get_intersection(&self, row_idx: usize, col_idx: usize) -> Option<&Data> {
        //Option()ならSome()が出てくるのかと思ってた。そんなこともないのか：俺
        self.rows.get(row_idx)?.get(col_idx)
    }
}

// ==========================================
// 【Step 2】聖域（main関数）
// ==========================================

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = "202606929中小企業実態基本調査_個人企業.xlsx";

    // 1. 関所を構築
    let scraper = ExcelScraper::new(path)?;
    println!("✅ Excelの読み込みと浄化に成功。総行数: {}", scraper.rows.len());

    // 2. クエリの定義
    let query = TargetQuery {
        row_name: "売上高".to_string(),
        col_name: "情報サービス業".to_string(),
    };

    // 3. 列と行の特定
    let col_idx = scraper.find_col(&query.col_name).ok_or("❌ 列が見つかりません")?;
    let row_idx = scraper.find_row(&query.row_name).ok_or("❌ 行が見つかりません")?;

    // 4. 交差点の抽出と出力の整形
    if let Some(value) = scraper.get_intersection(row_idx, col_idx) {
        println!("🎯 交差点の特定に成功！");
        println!("【{}】行インデックス: {}", query.row_name, row_idx);
        println!("【{}】列インデックス: {}", query.col_name, col_idx);
        
        // Claudeの指摘を反映し、生データ（Data::Float等）ではなく綺麗な数値として出力する
        //ここよくわからんな、なんで型をmatchする必要があるんだ？：俺
        match value {
            Data::Float(f) => println!("抽出された値: {}", f),
            Data::Int(i) => println!("抽出された値: {}", i),
            Data::String(s) => println!("抽出された値: {}", s),
            _ => println!("抽出された値: {:?}", value), // その他の型（空セルなど）の場合
        }
    } else {
        println!("❌ 指定された座標に値が存在しません。");
    }

    Ok(())
}
//P001_03_2_excel_セルのデータ抽出_test_AI生成_読解
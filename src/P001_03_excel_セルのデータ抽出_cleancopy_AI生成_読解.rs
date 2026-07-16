use calamine::{Data, DataType, Range, Reader, Xlsx, open_workbook};

/// 検索したい行と列の情報をまとめた構造体（名詞）
struct TargetQuery {
    row_name: String, // 例: "売上高"
    col_name: String, // 例: "情報サービス業"
}

/// 外界（Excelファイル）との境界を守り、データを提供する関所（Gateway）
/// Data は所有型なので、Range<Data> をそのまま持てばよく、
/// Vec<Vec<Data>> へのコピーは不要（12番目の設計を採用）
struct ExcelScraper {
    sheet: Range<Data>,
}

impl ExcelScraper {
    /// 新しい関所を構築する
    fn new(path: &str) -> Result<Self, Box<dyn std::error::Error>> {
        let mut workbook: Xlsx<_> = open_workbook(path)?;
        let range = workbook
            .worksheet_range_at(0)
            .ok_or("シートがありません")??;

        Ok(ExcelScraper { sheet: range })
    }

    /// 指定されたヘッダー名が存在する「列インデックス」を探す（先頭10行を横走査）
    fn find_column_index(&self, header_name: &str) -> Option<usize> {
        //take(10)を使うことで、for2回分を1回にまとめている：俺
        //for row_idx in 0..10 { take(10)とかくことで、row_idxを要らなくしている：俺
            //if let Some(row) = self.sheet.rows().nth(row_idx) {：俺
        for row in self.sheet.rows().take(10) {
            for (col_idx, cell) in row.iter().enumerate() {
                if let Some(text) = cell.as_string() {
                    if text.trim() == header_name {
                        return Some(col_idx);
                    }
                }
            }
        }
        None
    }

    /// 指定された項目名が存在する「行インデックス」を探す（各行の先頭3列を縦走査）
    fn find_row_index(&self, item_name: &str) -> Option<usize> {
        for (row_idx, row) in self.sheet.rows().enumerate() {
            for cell in row.iter().take(3) {
                if let Some(text) = cell.as_string() {
                    if text.trim() == item_name {
                        return Some(row_idx);
                    }
                }
            }
        }
        None
    }

    /// 指定した行と列のインデックスから、セルの値を取得する
    fn get_value_at(&self, row_idx: usize, col_idx: usize) -> Option<&Data> {
        self.sheet.get((row_idx, col_idx))
    }
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = "202606929中小企業実態基本調査_個人企業.xlsx";

    // 1. 関所の構築
    let scraper = ExcelScraper::new(path)?;
    println!(
        "✅ シートの読み込みに成功。総行数: {}, 総列数: {}",
        scraper.sheet.height(),
        scraper.sheet.width()
    );

    // 2. 抽出したいターゲット（クエリ）の定義
    let query = TargetQuery {
        row_name: "売上高".to_string(),
        col_name: "情報サービス業".to_string(),
    };

    // 3. 列と行のインデックス特定
    let col_idx = scraper
        .find_column_index(&query.col_name)
        .ok_or(format!("❌ 「{}」という業界が見つかりません", query.col_name))?;
    let row_idx = scraper
        .find_row_index(&query.row_name)
        .ok_or(format!("❌ 「{}」という項目が見つかりません", query.row_name))?;

    // 4. 交差点の抽出と、見やすい形式での出力
    match scraper.get_value_at(row_idx, col_idx) {
        Some(Data::Float(f)) => println!("🎯 抽出された値: {}", f),
        Some(Data::Int(i)) => println!("🎯 抽出された値: {}", i),
        Some(Data::String(s)) => println!("🎯 抽出された値: {}", s),
        Some(other) => println!("🎯 抽出された値（その他型）: {:?}", other),
        None => println!(
            "❌ 交差点 (行:{}, 列:{}) にデータが存在しませんでした。",
            row_idx, col_idx
        ),
    }

    Ok(())
}
//P001_03_excel_セルのデータ抽出_cleancopy_AI生成_読解
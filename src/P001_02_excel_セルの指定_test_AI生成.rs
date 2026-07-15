//----「：俺」は俺のコメント

//calamineのライブラリのコンパイラのヒントが凄く読みやすくてわかりやすい：俺
use calamine::{Reader, Xlsx, open_workbook, DataType};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = "202606929中小企業実態基本調査_個人企業.xlsx";
    let mut workbook: Xlsx<_> = open_workbook(path)?;

    // 1. 最初（インデックス0）のシートを取得する
    //ちゃんとエラーの条件分岐をするんだね。プログラムを組むのはきめ細かい作業だ：俺
    if let Some(Ok(range)) = workbook.worksheet_range_at(0) {
        //height()、width()みたいなcalamine由来の機能はわかりやすくていいね：俺
        println!("シートの読み込みに成功。総行数: {}, 総列数: {}", range.height(), range.width());

        // 2. 「情報サービス業」が何列目にあるか探索する
        // 今回の表は複数行にわたってヘッダー（業界名）が書かれている可能性があるため、
        // 最初の10行分を対象に横走査（ループ）をかけます。
        let mut target_col_index = None;

        //「'outer」←なんだこれ？ライフタイムか？：俺
        'outer: for row_idx in 0..10 {
            //Some()も時々見かけるけど、なんだこれ？：俺
            if let Some(row) = range.rows().nth(row_idx) {
                //rowの要素は&Dataなのえらい。enumerate()でcol_idxがusizeになっている：俺
                for (col_idx, cell) in row.iter().enumerate() {
                    // セルが文字列型で、かつ「情報サービス業」という文字が含まれているかチェック
                    //as_string()とto_string()の違いがよくわからなったけど、as_string()はdata typeをStringに直すようだ：俺
                    if let Some(text) = cell.as_string() {
                        if text.trim() == "情報サービス業" {
                            target_col_index = Some(col_idx);
                            println!("🎯 見つけました！");
                            println!("「情報サービス業」は {} 行目の {} 列目（インデックス: {}）にあります。", row_idx + 1, col_idx + 1, col_idx);
                            break 'outer; // 見つかったら外側のループごと抜ける
                        }
                    }
                }
            }
        }

        if target_col_index.is_none() {
            println!("❌ 「情報サービス業」という項目が見つかりませんでした。");
        }

    } else {
        println!("❌ シートが見つかりませんでした。");
    }

    //Ok()ってなんだろう？main()からResultを外に返しているのかな？：俺
    Ok(())
}
//P001_02_excel_セルの指定_test_AI生成_読解
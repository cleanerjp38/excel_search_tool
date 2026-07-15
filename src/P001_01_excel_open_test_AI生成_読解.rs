//calamineはExcelに関する定番ライブラリらしい
//ライブラリってなにが入ってるものなんだろ？
use calamine::{Reader, Xlsx, open_workbook};

//おお！？main()が何かを返しているぞ？こんなのは初めて見た
//Result<(), Box<dyn std::error::Error>> なんだこれは。Boxとdynは以前に一度見たが、謎のままだった
//前にも思ったが、BoxはVecに書き方が似ている。何かが可変なのかな？
fn main() -> Result<(), Box<dyn std::error::Error>> {
    // 1. Excelファイルのパスを指定して開く
    let path = "202606929中小企業実態基本調査_個人企業.xlsx";
    println!("Excelファイルを読み込み中: {}...", path);
    
    //Xlsxって型があるのか！これがExcelを指定しているんだな
    //open_workbook()はライブラリ内に準備されている機能だろう
    //「?」はどんな意味を持っているんだ？
    let mut workbook: Xlsx<_> = open_workbook(path)?;

    // 2. ファイルに含まれるシート名の一覧を取得して表示してみる
    let sheet_names = workbook.sheet_names();
    println!("正常に開けました！");
    println!("シート一覧: {:?}", sheet_names);

    //なんだこれ？
    Ok(())
}
//P001_01_excel_open_test_AI生成_読解
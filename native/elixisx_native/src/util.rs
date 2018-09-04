pub fn to_excel_coords(y: i32, x: i32) -> String {
    encode_col(x) + y.to_string().as_str()
}

pub fn encode_col(col: i32) -> String {
    match col {
        0 => "".to_string(),
        x if x <= 26 => ((x + 64) as u8 as char).to_string(),
        x => {
            let m = x / 26;
            let r = x % 26;
            if r == 0{
                encode_col(m - 1) + encode_col(26).as_str()
            }else{
                encode_col(m) + encode_col(r).as_str()
            }
        }
    }
}

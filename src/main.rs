use actix_web::{get, post, web, App, HttpResponse, HttpServer, Responder, Result};
use serde::{Serialize, Deserialize};
use xlsxwriter::*;
use actix_files::NamedFile;
use std::path::PathBuf;

use actix_web::http::ContentEncoding;
use actix_web::dev::BodyEncoding;
use std::io::Read;
use std::path::Path;
use std::fs::File;
use uuid::Uuid;
use std::fs;
//use actix_cors::Cors;

#[derive(Serialize, Deserialize, Debug)]
struct BaseInitExcel {
   name: String, 
   lists: Vec<BaseInitList>
}


#[derive(Serialize, Deserialize, Debug)]
struct BaseInitList {
    name: String,
    params: Vec<BaseInitСell>
}

#[derive(Serialize, Deserialize, Debug)]
struct BaseInitСell {
    coord: String,
    param: BaseInitParam
}

#[derive(Serialize, Deserialize, Debug)]
struct BaseInitParam {
    caption: String,
    style: Option<String>,//тут будет объект
    merge_sell: Option<String>,
    format: Option<String>,
    formula: Option<bool>,
    num_format:Option<String>
}

struct BaseCoordExcel{
    column:i32,
    row:i32
}
struct BaseCoordBetween{
    start:BaseCoordExcel,
    end:BaseCoordExcel
}

fn intoCoord(_coord: String) -> BaseCoordBetween{
    let chars: Vec<_> = _coord.chars().collect();
    let mut t:i32 = 0;
    let mut n:i32 = 0;
    let mut a:i32 = 0;
    let mut tmp:i32 = -1;
    let ln:usize = _coord.len();
    for ch in _coord.chars().collect::<Vec<char>>() {
        tmp +=1;
        if tmp < a{
            continue;
        }
        n = ch as i32;
        n = n - 64;
        if (n  < 1) || n > 26{
            break;
        }
        t = 26 * t + n;
        a +=1;
    }
    let column:i32 = t - 1;
    tmp = -1;
    t = 0;
    for ch in _coord.chars().collect::<Vec<char>>() {
        tmp +=1;
        if tmp < a{
            continue;
        }
        n = ch as i32;
        n = n - 48;
        if (n < 0) || n > 9{
            break;
        }
        t = 10 * t + n;
        a +=1;
    }
    let row:i32 = t - 1;
    let mut baseCoordStart = BaseCoordExcel{
        column: column,
        row: row
    };
    let baseCoordEnd = BaseCoordExcel{
        column: -1,
        row: -1
    };

    let mut baseCoord = BaseCoordBetween{
        start:baseCoordStart,
        end: baseCoordEnd
    };
    if(a < ln.try_into().unwrap()){
        t = 0;
        tmp = -1;
        for ch in _coord.chars().collect::<Vec<char>>() {
            tmp +=1;
            if tmp <= a{
                continue;
            }
            n = ch as i32;
            n = n - 64;
            if (n  < 1) || n > 26{
                break;
            }
            t = 26 * t + n;
            a +=1;
        }
        let column:i32 = t - 1;
        baseCoord.end.column =  column;
        tmp = -1;
        t = 0;
        for ch in _coord.chars().collect::<Vec<char>>() {
            tmp +=1;
            if tmp <= a{
                continue;
            }
            n = ch as i32;
            n = n - 48;
            if (n < 0) || n > 9{
                break;
            }
            t = 10 * t + n;
            a +=1;
        }
        baseCoord.end.row=  t - 1;
    }
    
    baseCoord
    
 }
 /*
 #[get("/")]
 async fn index(req_body: String) -> HttpResponse {
    intoCoord(String::from("A2"));
    let workbook = Workbook::new("simple1.xlsx");
    let mut format1 = workbook.add_format()
         .set_font_color(FormatColor::Red);
 
    let mut format2 = workbook.add_format()
         .set_font_color(FormatColor::Blue)
         .set_underline(FormatUnderline::Single);
 
    let mut format3 = workbook.add_format()
         .set_font_color(FormatColor::Green)
         .set_align(FormatAlignment::CenterAcross)
         .set_align(FormatAlignment::VerticalCenter);
 
    let mut sheet1 = workbook.add_worksheet(Some("привет")).expect("oh no! function() failed!!");
    sheet1.write_string(0, 0, "Red text", Some(&format1)).expect("oh no! function() failed!!");
    sheet1.write_number(0, 1, 20., None).expect("oh no! function() failed!!");
    sheet1.write_formula_num(1, 0, "=10+B1", None, 30.).expect("oh no! function() failed!!");
    sheet1.write_url(
         1,
         1,
         "https://github.com/informationsea/xlsxwriter-rs",
         Some(&format2),
     ).expect("oh no! function() failed!!");
     sheet1.merge_range(2, 0, 3, 2, "Hello, world", Some(&format3)).expect("oh no! function() failed!!");
 
     sheet1.set_selection(1, 0, 1, 2);
     sheet1.set_tab_color(FormatColor::Cyan);
     workbook.close().expect("oh no! function() failed!!");
    //let path: PathBuf = "simple1.xlsx".parse().unwrap();
   //  Ok(NamedFile::open(path).expect("oh no! function() failed!!"))
    let mut f = File::open("simple1.xlsx").unwrap();
    let mut buffer = Vec::new();

    // read the whole file
    f.read_to_end(&mut buffer);

    HttpResponse::Ok()
        .encoding(ContentEncoding::Identity)
        .content_type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        .header("accept-ranges", "bytes")
        .header("content-disposition", "attachment; filename=\"tttt.xlsx\"")
        .body(buffer)
 }
*/
#[post("/export/v1/")]
async fn exportV1(req_body: String) -> impl Responder /*Result<NamedFile>*/ {
    let v: BaseInitExcel = serde_json::from_str(&req_body).expect("oh no! function() failed!!");
    let my_uuid = Uuid::new_v4();
    let lists: Vec<BaseInitList> = v.lists;
   // let workbook = Workbook::new(&v.name);
    let workbook = Workbook::new(&my_uuid.to_string());
    for item in &lists {
        let mut sheet = workbook.add_worksheet(Some(&item.name)).expect("oh no! function() failed!!");
        for params in &item.params {
            let coord = &params.coord;//координаты ячеек A1;
            let t = intoCoord(String::from(&params.coord));
            let param  = &params.param;//координаты ячеек A1;

            let format = match &param.format {
                None => String::from(""),
                Some(x) => String::from(x) 
            };
            
            let formula = match &param.formula{
                None => false,
                Some(x) => *x
            };
            let num_format = match &param.num_format{
                None => String::from(""),
                Some(x) => String::from(x)
            };
            
            let mergeSell = match &param.merge_sell {
                None => String::from(""),
                Some(x) => String::from(x) 
            };
            let column = (*&t.start.column as i32);
            let row = (*&t.start.row as u16);
            let mut option = workbook.add_format();
            if num_format != ""{
                option = option.set_num_format(&num_format);
            }

            if mergeSell != ""{
                let mut mergCoord: String = (coord).to_owned();
                mergCoord.push_str(":");
                mergCoord.push_str(&mergeSell);
                let mCoord: BaseCoordBetween = intoCoord(String::from(&mergCoord));

                let columnStart = (*&mCoord.start.column as i32);
                let rowStart = (*&mCoord.start.row as u16);
                let columnEnd = (*&mCoord.end.column as i32);
                let rowEnd = (*&mCoord.end.row as u16);
                println!("{}", &param.caption);
                sheet.merge_range(columnStart.try_into().unwrap(), rowStart, columnEnd.try_into().unwrap(), rowEnd, &param.caption, None).expect("oh no! function() failed!!");
                /*if formula == true {
                    sheet.write_formula(columnStart.try_into().unwrap(), rowStart, &param.caption, Some(&option)).expect("oh no! function() failed!!");
                }*/
            }
            if format == "string"{
                sheet.write_string(column.try_into().unwrap(), row, &param.caption, None).expect("oh no! function() failed!!");
            }else if format == "number"{
                let caption:f64 = (&param.caption).parse().unwrap();
                sheet.write_number(column.try_into().unwrap(), row, caption, None).expect("oh no! function() failed!!");
            }
            if mergeSell == ""{
                         
                sheet.write_string(column.try_into().unwrap(), row, &param.caption,  Some(&option)).expect("oh no! function() failed!!");
                /*if formula == true {
                    sheet.write_formula(column.try_into().unwrap(), row, &param.caption, Some(&option)).expect("oh no! function() failed!!");
                }*/
                /*if format == "string"{
                    sheet.write_string(column.try_into().unwrap(), row, &param.caption, None).expect("oh no! function() failed!!");
                }else if format == "number"{
                    let caption:f64 = (&param.caption).parse().unwrap();
                    sheet.write_number(column.try_into().unwrap(), row, caption, None).expect("oh no! function() failed!!");
                }*/
            }
        }
    }
    workbook.close().expect("oh no! function() failed!!");


    let mut f = File::open(&my_uuid.to_string()).unwrap();
    let mut buffer = Vec::new();

    // read the whole file
    f.read_to_end(&mut buffer);
    let mut fl_name: String = ("attachment; filename=\"").to_owned();
    fl_name.push_str(&v.name);
    fl_name.push_str("\"");
    fs::remove_file(&my_uuid.to_string());
    HttpResponse::Ok()
        .encoding(ContentEncoding::Identity)
        .content_type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        .header("accept-ranges", "bytes")
        .header("content-disposition", fl_name)
        .body(buffer)
}


#[actix_web::main]
async fn main() -> std::io::Result<()> {
    HttpServer::new(|| {
        App::new()
            .service(exportV1)
    })
    .bind(("127.0.0.1", 8081))?
    .run()
    .await
}
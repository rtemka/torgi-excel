use calamine::{open_workbook, DataType, Range, Reader, Xlsx, XlsxError};
use serde::{Deserialize, Serialize};
use std::{fmt, fs::File, io::BufReader, path::Path};

const DATE_ENDING_BIDS: &str = "Дата_окончания_подачи_заявок";
const TIME_ENDING_BIDS: &str = "Время_окончания_подачи_заявок";
const DATE_APPROVAL: &str = "Дата_окончания_срока_рассмотрения_заявок";
const DATE_BIDDING: &str = "Дата_проведения_аукциона_конкурса";
const TIME_BIDDING: &str = "Время_проведения_аукциона_конкурса";
const PURCHASE_SUBJECT: &str = "Предмет";
const REGION: &str = "Регион";
const CUSTOMER_TYPE: &str = "Заказчик";
const MAX_PRICE: &str = "НМЦК";
const NUMBER: &str = "Номер";
const WINNER: &str = "Победитель";
const STATUS: &str = "Статус";
const ESTIMATION: &str = "Расчет";
const WINNER_PRICE: &str = "Сумма_выигранного_лота";
const PARTICIPANTS: &str = "Участники";
const PURCHASE_TYPE: &str = "Форма_проведения";
const EMPTY: &str = "";
const STATUS_GO: &str = "идем";
const STATUS_ADMITTED: &str = "допущены";
const STATUS_APPLY: &str = "заявлены";
const NAMED_RANGES_COUNT: usize = 16;
const TOTAL_COLUMNS: u32 = 16384;
const RADIX: u32 = 36;
const WORKBOOK_MAX_ROWS: usize = 7000;

#[cfg(test)]
pub fn print_active_state_json() -> Result<(), WorkbookError> {
    let mut workbook = open_wb()?;

    let named_ranges = named_ranges(&workbook);

    let named_cols = named_cols(&named_ranges);

    if let Some(Ok(range)) = workbook.worksheet_range(&named_ranges[0].sheet) {
        let purches = active_state_cells(range, named_cols);
        for p in purches {
            let j = serde_json::to_string(&p).map_err(|_| WorkbookError::JsonSerializeError)?;
            println!("{}", &j);
        }
    }
    Ok(())
}

#[cfg(test)]
pub fn print_active_state() -> Result<(), WorkbookError> {
    let mut workbook = open_wb()?;

    let named_ranges = named_ranges(&workbook);

    let named_cols = named_cols(&named_ranges);

    if let Some(Ok(range)) = workbook.worksheet_range(&named_ranges[0].sheet) {
        let purches = active_state_cells(range, named_cols);
        for p in purches {
            println!(
                "rn:{:?} ps:{:?} reg:{:?} ct:{:?} status:{:?} est:{:?} bid_date:{:?} bid_time{:?}",
                p.registry_number,
                p.purchase_subject,
                p.region,
                p.customer_type,
                p.status,
                p.estimation,
                p.bidding_date,
                p.bidding_time
            );
        }
    }
    Ok(())
}

pub fn get_active_state_json() -> Result<String, WorkbookError> {
    let mut workbook = open_wb()?;

    let named_ranges = named_ranges(&workbook);

    let named_cols = named_cols(&named_ranges);

    let mut json = String::new();

    if let Some(Ok(range)) = workbook.worksheet_range(&named_ranges[0].sheet) {
        let purches = active_state_cells(range, named_cols);

        json = serde_json::to_string(&purches).map_err(|_| WorkbookError::JsonSerializeError)?;
    }
    Ok(json)
}

#[cfg(test)]
pub fn get_active_state() -> Result<Vec<Purchase>, WorkbookError> {
    let mut workbook = open_wb()?;

    let named_ranges = named_ranges(&workbook);

    let named_cols = named_cols(&named_ranges);

    let mut purches: Vec<Purchase> = Vec::new();

    if let Some(Ok(range)) = workbook.worksheet_range(&named_ranges[0].sheet) {
        purches = active_state_cells(range, named_cols);
    }
    Ok(purches)
}

fn open_wb() -> Result<Xlsx<BufReader<File>>, WorkbookError> {
    let path = Path::new(crate::WORKBOOK_PATH);
    Ok(open_workbook(path).map_err(|x| WorkbookError::XlsxError(x))?)
}

#[derive(Debug)]
pub enum WorkbookError {
    InvalidColumnNameError(String),
    XlsxError(XlsxError),
    JsonSerializeError,
    #[allow(dead_code)]
    IoError(std::io::Error),
}

enum ColumnPosition {
    Left = 0,
    #[allow(dead_code)]
    Right = 1,
}

impl fmt::Display for WorkbookError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self {
            WorkbookError::InvalidColumnNameError(s) => write!(f, "invalid column name: {:?}", &s),
            WorkbookError::XlsxError(e) => write!(f, "{:?}", &e),
            WorkbookError::JsonSerializeError => write!(f, "cannot serialize struct"),
            WorkbookError::IoError(e) => write!(f, "{:?}", &e),
        }
    }
}

#[derive(Debug)]
struct NamedCols {
    registry_number_col: usize,
    purchase_subject_col: usize,
    customer_type_col: usize,
    status_col: usize,
    region_col: usize,
    estimation_col: usize,
    bidding_date_col: usize,
    bidding_time_col: usize,
}

#[derive(Serialize, Deserialize, Debug)]
pub struct Purchase {
    registry_number: String,
    purchase_subject: String,
    region: String,
    status: String,
    customer_type: String,
    estimation: f64,
    bidding_date: f64,
    bidding_time: f64,
}

struct NamedRange<'a> {
    name: &'a str,
    sheet: String,
    range: String,
}

impl NamedRange<'_> {
    pub fn new(def_name: &(String, String)) -> Self {
        Self {
            name: Self::match_name(&def_name.0),
            sheet: Self::parse_sheet(&def_name.1),
            range: Self::parse_range(&def_name.1),
        }
    }

    fn match_name<'a>(s: &str) -> &'a str {
        match s {
            DATE_ENDING_BIDS => DATE_ENDING_BIDS,
            TIME_ENDING_BIDS => TIME_ENDING_BIDS,
            DATE_APPROVAL => DATE_APPROVAL,
            DATE_BIDDING => DATE_BIDDING,
            TIME_BIDDING => TIME_BIDDING,
            MAX_PRICE => MAX_PRICE,
            NUMBER => NUMBER,
            WINNER => WINNER,
            STATUS => STATUS,
            WINNER_PRICE => WINNER_PRICE,
            PARTICIPANTS => PARTICIPANTS,
            PURCHASE_TYPE => PURCHASE_TYPE,
            PURCHASE_SUBJECT => PURCHASE_SUBJECT,
            CUSTOMER_TYPE => CUSTOMER_TYPE,
            ESTIMATION => ESTIMATION,
            REGION => REGION,
            _ => EMPTY,
        }
    }

    fn parse_sheet(s: &str) -> String {
        s.split('!').nth(0).unwrap_or_default().to_owned()
    }

    fn parse_range(s: &str) -> String {
        s.split('!').nth(1).unwrap_or_default().to_owned()
    }

    pub fn col_name(&self, cp: ColumnPosition) -> String {
        self.range
            .split(":")
            .nth(cp as usize)
            .unwrap_or_default()
            .replace("$", "")
    }

    pub fn col_num(&self, cp: ColumnPosition) -> Result<usize, WorkbookError> {
        let col_name = self.col_name(cp);

        if col_name.len() == 0 {
            return Err(WorkbookError::InvalidColumnNameError(col_name));
        }

        let mut col = 0;
        let mut multi = 1;

        for c in col_name.chars().rev() {
            match c {
                'A'..='Z' => {
                    if let (Some(c_num), Some(a_num)) = (c.to_digit(RADIX), 'A'.to_digit(RADIX)) {
                        col += (c_num - a_num + 1) * multi;
                    }
                }
                'a'..='z' => {
                    if let (Some(c_num), Some(a_num)) = (c.to_digit(RADIX), 'a'.to_digit(RADIX)) {
                        col += (c_num - a_num + 1) * multi;
                    }
                }
                _ => return Err(WorkbookError::InvalidColumnNameError(col_name)),
            }
            multi *= 26;
        }

        if col > TOTAL_COLUMNS {
            return Err(WorkbookError::InvalidColumnNameError(col_name));
        }

        Ok(col as usize - 1)
    }
}

fn named_ranges<'a>(workbook: &Xlsx<BufReader<File>>) -> Vec<NamedRange<'a>> {
    let mut named_ranges: Vec<NamedRange> = Vec::with_capacity(NAMED_RANGES_COUNT);

    for name in workbook
        .defined_names()
        .iter()
        .filter(|(n, _)| is_expectable(n))
    {
        named_ranges.push(NamedRange::new(name));
    }

    named_ranges
}

fn named_cols(named_ranges: &Vec<NamedRange>) -> NamedCols {
    let mut named_cols = NamedCols {
        registry_number_col: 0,
        purchase_subject_col: 0,
        customer_type_col: 0,
        status_col: 0,
        region_col: 0,
        estimation_col: 0,
        bidding_date_col: 0,
        bidding_time_col: 0,
    };

    for nr in named_ranges {
        let col = nr.col_num(ColumnPosition::Left).unwrap_or_default();
        match nr.name {
            DATE_BIDDING => named_cols.bidding_date_col = col,
            TIME_BIDDING => named_cols.bidding_time_col = col,
            NUMBER => named_cols.registry_number_col = col,
            STATUS => named_cols.status_col = col,
            PURCHASE_SUBJECT => named_cols.purchase_subject_col = col,
            CUSTOMER_TYPE => named_cols.customer_type_col = col,
            ESTIMATION => named_cols.estimation_col = col,
            REGION => named_cols.region_col = col,
            _ => continue,
        }
    }

    named_cols
}

fn active_state_cells(rng: Range<DataType>, cols: NamedCols) -> Vec<Purchase> {
    let mut active_cells: Vec<Purchase> = Vec::new();

    let result = rng
        .rows()
        .take(WORKBOOK_MAX_ROWS)
        .filter(|r| match &r[cols.status_col] {
            DataType::String(s) => is_active_state(&s),
            _ => false,
        });

    for r in result {
        let purch = Purchase {
            registry_number: r[cols.registry_number_col]
                .get_string()
                .unwrap_or_default()
                .replace("№", ""),
            purchase_subject: r[cols.purchase_subject_col]
                .get_string()
                .unwrap_or_default()
                .to_owned(),
            region: r[cols.region_col]
                .get_string()
                .unwrap_or_default()
                .to_owned(),
            customer_type: r[cols.customer_type_col]
                .get_string()
                .unwrap_or_default()
                .to_owned(),
            estimation: r[cols.estimation_col].get_float().unwrap_or_default(),
            bidding_date: match r[cols.bidding_date_col] {
                DataType::DateTime(dt) => dt,
                _ => 0.0,
            },
            bidding_time: r[cols.bidding_time_col].get_float().unwrap_or_default(),
            status: r[cols.status_col]
                .get_string()
                .unwrap_or_default()
                .to_owned(),
        };

        active_cells.push(purch);
    }

    active_cells
}

fn is_expectable(s: &str) -> bool {
    match s {
        DATE_ENDING_BIDS => true,
        TIME_ENDING_BIDS => true,
        DATE_APPROVAL => true,
        DATE_BIDDING => true,
        TIME_BIDDING => true,
        MAX_PRICE => true,
        NUMBER => true,
        WINNER => true,
        STATUS => true,
        WINNER_PRICE => true,
        PARTICIPANTS => true,
        PURCHASE_TYPE => true,
        PURCHASE_SUBJECT => true,
        CUSTOMER_TYPE => true,
        ESTIMATION => true,
        REGION => true,
        _ => false,
    }
}

fn is_active_state(s: &str) -> bool {
    match s {
        STATUS_GO => true,
        STATUS_ADMITTED => true,
        STATUS_APPLY => true,
        _ => false,
    }
}

mod tests {

    #[test]
    fn test_print_active_state() {
        assert!(super::print_active_state().is_ok());
    }

    #[test]
    fn test_print_active_state_json() {
        assert!(super::print_active_state_json().is_ok());
    }

    #[test]
    fn test_get_active_state_json() {
        assert!(super::get_active_state_json().is_ok());
    }

    #[test]
    fn test_get_active_state() {
        assert!(super::get_active_state().is_ok());
    }
}

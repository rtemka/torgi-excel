use crate::simple_time::Moment;
use calamine::{open_workbook, DataType, Range, Reader, Xlsx, XlsxError};
use serde::{Deserialize, Serialize};
use std::{
    fmt,
    fs::File,
    io::BufReader,
    path::Path,
    time::{Duration, SystemTime},
};

const PURCHASE_SUBJECT_ABBR: &str = "Предмет";
const PURCHASE_SUBJECT: &str = "Поставляемые_товары";
const DATE_COLLECTING_BIDS: &str = "Дата_окончания_подачи_заявок";
const TIME_COLLECTING_BIDS: &str = "Время_окончания_подачи_заявок";
const DATE_APPROVAL: &str = "Дата_окончания_срока_рассмотрения_заявок";
const DATE_BIDDING: &str = "Дата_проведения_аукциона_конкурса";
const TIME_BIDDING: &str = "Время_проведения_аукциона_конкурса";
const REGION: &str = "Регион";
const CUSTOMER_TYPE: &str = "Заказчик";
const MAX_PRICE: &str = "НМЦК";
const APPLICATION_GUARANTEE: &str = "Размер_обеспечения_заявки";
const CONTRACT_GUARANTEE: &str = "Размер_обеспечения_контракта";
const STATUS: &str = "Статус";
const OUR_PARTICIPANTS: &str = "Наши_участники";
const ESTIMATION: &str = "Расчет";
const NUMBER: &str = "Номер";
const PURCHASE_TYPE: &str = "Форма_проведения";
const ETP: &str = "Площадка";
const WINNER: &str = "Победитель";
const WINNER_PRICE: &str = "Сумма_выигранного_лота";
const PARTICIPANTS: &str = "Участники";
const EMPTY: &str = "";
const STATUS_GO: &str = "идем";
const STATUS_ADMITTED: &str = "допущены";
const STATUS_APPLY: &str = "заявлены";
const STATUS_WIN: &str = "выиграли";
const STATUS_LOSS: &str = "не выиграли";
const NAMED_RANGES_COUNT: usize = 20;
const TOTAL_COLUMNS: u32 = 16384;
const RADIX: u32 = 36;
const WORKBOOK_MAX_ROWS: usize = 7000;
const EXCEL_UNIX_EPOCH: u64 = 25569;
const HOW_FAR_IN_PAST_DAYS: u64 = 16;
const FILE_UTC_TIME_OFFSET: &str = "+03:00";

#[cfg(test)]
pub fn print_active_state_json() -> Result<(), WorkbookError> {
    let mut workbook = open_wb()?;

    if let Some(purches) = active_purchases(&mut workbook) {
        println!("{}", purches.len());
        for p in purches {
            let j =
                serde_json::to_string_pretty(&p).map_err(|_| WorkbookError::JsonSerializeError)?;
            println!("{}", &j);
        }
    }
    Ok(())
}

#[cfg(test)]
pub fn print_active_state() -> Result<(), WorkbookError> {
    let mut workbook = open_wb()?;

    if let Some(purches) = active_purchases(&mut workbook) {
        for p in purches {
            println!(
                "rn:{:?} ps:{:?} reg:{:?} ct:{:?} status:{:?} est:{:?} bid_datetime:{:?}",
                p.registry_number,
                p.purchase_subject,
                p.region,
                p.customer_type,
                p.status,
                p.estimation,
                p.bidding_datetime,
            );
        }
    }

    Ok(())
}

pub fn get_active_state_json() -> Result<Option<String>, WorkbookError> {
    let mut workbook = open_wb()?;

    if let Some(purches) = active_purchases(&mut workbook) {
        Ok(Some(
            serde_json::to_string(&purches).map_err(|_| WorkbookError::JsonSerializeError)?,
        ))
    } else {
        Ok(None)
    }
}

#[cfg(test)]
pub fn get_active_state() -> Result<Option<Vec<Purchase>>, WorkbookError> {
    let mut workbook = open_wb()?;

    Ok(active_purchases(&mut workbook))
}

fn open_wb() -> Result<Xlsx<BufReader<File>>, WorkbookError> {
    let path = Path::new(crate::WORKBOOK_PATH);
    Ok(open_workbook(path).map_err(|x| WorkbookError::XlsxError(x))?)
}

fn active_purchases(workbook: &mut Xlsx<BufReader<File>>) -> Option<Vec<Purchase>> {
    let named_ranges = named_ranges(&workbook);

    let named_cols = named_cols(&named_ranges);

    match workbook.worksheet_range(&named_ranges[0].sheet) {
        Some(Ok(range)) => {
            let purches = active_state_cells(range, named_cols);
            if purches.is_empty() {
                None
            } else {
                Some(purches)
            }
        }
        _ => None,
    }
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

#[derive(Default, Debug)]
struct NamedCols {
    registry_number: usize,
    purchase_subject: usize,
    purchase_abbr: usize,
    purchase_type: usize,
    collecting_datetime: usize,
    approval_datetime: usize,
    bidding_datetime: usize,
    bidding_date: usize,
    collecting_date: usize,
    region: usize,
    customer_type: usize,
    max_price: usize,
    application_guarantee: usize,
    contract_guarantee: usize,
    status: usize,
    our_participants: usize,
    estimation: usize,
    etp: usize,
    winner: usize,
    winner_price: usize,
    participants: usize,
}

#[derive(Serialize, Deserialize, Debug)]
pub struct Purchase {
    registry_number: String,
    purchase_subject: String,
    purchase_abbr: String,
    purchase_type: String,
    collecting_datetime: String,
    approval_datetime: String,
    bidding_datetime: String,
    region: String,
    customer_type: String,
    max_price: f64,
    application_guarantee: f64,
    contract_guarantee: f64,
    status: String,
    our_participants: String,
    estimation: f64,
    etp: String,
    winner: String,
    winner_price: f64,
    participants: String,
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
            DATE_COLLECTING_BIDS => DATE_COLLECTING_BIDS,
            TIME_COLLECTING_BIDS => TIME_COLLECTING_BIDS,
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
            PURCHASE_SUBJECT_ABBR => PURCHASE_SUBJECT_ABBR,
            OUR_PARTICIPANTS => OUR_PARTICIPANTS,
            ETP => ETP,
            APPLICATION_GUARANTEE => APPLICATION_GUARANTEE,
            CONTRACT_GUARANTEE => CONTRACT_GUARANTEE,
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
    let mut named_cols = NamedCols::default();

    for nr in named_ranges {
        let col = nr.col_num(ColumnPosition::Left).unwrap_or_default();
        match nr.name {
            DATE_BIDDING => named_cols.bidding_date = col,
            DATE_COLLECTING_BIDS => named_cols.collecting_date = col,
            TIME_COLLECTING_BIDS => named_cols.collecting_datetime = col,
            DATE_APPROVAL => named_cols.approval_datetime = col,
            TIME_BIDDING => named_cols.bidding_datetime = col,
            NUMBER => named_cols.registry_number = col,
            STATUS => named_cols.status = col,
            PURCHASE_SUBJECT => named_cols.purchase_subject = col,
            CUSTOMER_TYPE => named_cols.customer_type = col,
            ESTIMATION => named_cols.estimation = col,
            REGION => named_cols.region = col,
            MAX_PRICE => named_cols.max_price = col,
            WINNER => named_cols.winner = col,
            WINNER_PRICE => named_cols.winner_price = col,
            PARTICIPANTS => named_cols.participants = col,
            PURCHASE_TYPE => named_cols.purchase_type = col,
            PURCHASE_SUBJECT_ABBR => named_cols.purchase_abbr = col,
            OUR_PARTICIPANTS => named_cols.our_participants = col,
            ETP => named_cols.etp = col,
            APPLICATION_GUARANTEE => named_cols.application_guarantee = col,
            CONTRACT_GUARANTEE => named_cols.contract_guarantee = col,
            _ => continue,
        }
    }

    named_cols
}

fn active_state_cells(rng: Range<DataType>, cols: NamedCols) -> Vec<Purchase> {
    let cut_off_date = today_in_excel_date() - HOW_FAR_IN_PAST_DAYS as f64;
    let rows = rng.rows().take(WORKBOOK_MAX_ROWS).filter(|r| {
        match (&r[cols.status], &r[cols.bidding_date]) {
            (DataType::String(s), DataType::DateTime(dt)) => {
                dt > &cut_off_date && is_active_state(&s)
            }
            _ => false,
        }
    });
    let mut purches: Vec<Purchase> = Vec::new();
    for r in rows {
        let mut bid_datetime = r[cols.bidding_datetime].get_float().unwrap_or_default();
        let bid_date = r[cols.bidding_date].get_float().unwrap_or_default();
        let mut col_datetime = r[cols.collecting_datetime].get_float().unwrap_or_default();
        let col_date = r[cols.collecting_date].get_float().unwrap_or_default();
        if bid_datetime < bid_date {
            bid_datetime += bid_date
        }
        if col_datetime < col_date {
            col_datetime += col_date
        }
        purches.push(Purchase {
            registry_number: r[cols.registry_number]
                .get_string()
                .unwrap_or_default()
                .replace("№", ""),
            purchase_subject: r[cols.purchase_subject]
                .get_string()
                .unwrap_or_default()
                .to_owned(),
            purchase_abbr: r[cols.purchase_abbr]
                .get_string()
                .unwrap_or_default()
                .to_owned(),
            purchase_type: r[cols.purchase_type]
                .get_string()
                .unwrap_or_default()
                .to_owned(),
            region: r[cols.region].get_string().unwrap_or_default().to_owned(),
            customer_type: r[cols.customer_type]
                .get_string()
                .unwrap_or_default()
                .to_owned(),
            max_price: r[cols.max_price].get_float().unwrap_or_default(),
            application_guarantee: r[cols.application_guarantee]
                .get_float()
                .unwrap_or_default(),
            contract_guarantee: r[cols.contract_guarantee].get_float().unwrap_or_default(),
            estimation: r[cols.estimation].get_float().unwrap_or_default(),
            our_participants: r[cols.our_participants]
                .get_string()
                .unwrap_or_default()
                .to_owned(),
            etp: r[cols.etp].get_string().unwrap_or_default().to_owned(),
            winner: r[cols.winner].get_string().unwrap_or_default().to_owned(),
            winner_price: r[cols.winner_price].get_float().unwrap_or_default(),
            participants: r[cols.participants]
                .get_string()
                .unwrap_or_default()
                .to_owned(),
            bidding_datetime: from_excel_date(bid_datetime),
            collecting_datetime: from_excel_date(col_datetime),
            approval_datetime: from_excel_date(match r[cols.approval_datetime] {
                DataType::DateTime(dt) => dt,
                _ => 0.0,
            }),
            status: r[cols.status].get_string().unwrap_or_default().to_owned(),
        });
    }
    purches
}

fn is_expectable(s: &str) -> bool {
    match s {
        DATE_COLLECTING_BIDS => true,
        TIME_COLLECTING_BIDS => true,
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
        PURCHASE_SUBJECT_ABBR => true,
        OUR_PARTICIPANTS => true,
        ETP => true,
        APPLICATION_GUARANTEE => true,
        CONTRACT_GUARANTEE => true,
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
        STATUS_WIN => true,
        STATUS_LOSS => true,
        _ => false,
    }
}

fn today_in_excel_date() -> f64 {
    if let Ok(duration) = SystemTime::now().duration_since(SystemTime::UNIX_EPOCH) {
        let d_since_epoch = days_from_seconds(duration.as_secs());
        let today_excel = d_since_epoch + EXCEL_UNIX_EPOCH;
        today_excel as f64
    } else {
        0.0
    }
}

fn to_rfc3339_string(m: Moment) -> String {
    // "2006-01-02T15:04:05Z07:00"
    format!(
        "{}-{}-{}T{}:{}:00{}",
        m.year,
        add_leading_zero(m.month),
        add_leading_zero(m.day),
        add_leading_zero(m.hours),
        add_leading_zero(m.minutes),
        FILE_UTC_TIME_OFFSET
    )
}

fn add_leading_zero(x: u64) -> String {
    if x < 10 {
        return format!("0{}", x);
    }
    x.to_string()
}

fn from_excel_date(excel_date: f64) -> String {
    let days = excel_date - EXCEL_UNIX_EPOCH as f64;
    let seconds = seconds_from_days(days);
    let duration = Duration::new(seconds, 0);
    to_rfc3339_string(Moment::from_duration_since_epoch(duration))
}

const fn days_from_seconds(timestamp_secs: u64) -> u64 {
    timestamp_secs / 86400
}

fn seconds_from_days(days: f64) -> u64 {
    (days * 86400.0).round() as u64
}

#[cfg(test)]
mod tests {

    use super::*;

    #[test]
    fn test_print_active_state() {
        assert!(print_active_state().is_ok());
    }

    #[test]
    fn test_print_active_state_json() {
        assert!(print_active_state_json().is_ok());
    }

    #[test]
    fn test_get_active_state_json() {
        assert!(get_active_state_json().is_ok());
    }

    #[test]
    fn test_get_active_state() {
        assert!(get_active_state().is_ok());
    }

    #[test]
    fn test_open_wb() {
        assert!(open_wb().is_ok());
    }

    #[test]
    fn test_add_leading_zero() {
        assert_eq!(add_leading_zero(9), "09".to_string());
        assert_eq!(add_leading_zero(11), "11".to_string());
    }

    #[test]
    fn test_to_rfc3339_string() {
        let m = Moment {
            year: 2021,
            month: 11,
            day: 10,
            hours: 12,
            minutes: 1,
            seconds: 44,
            is_leap_year: false,
        };
        assert_eq!(
            to_rfc3339_string(m),
            "2021-11-10T12:01:00+03:00".to_string()
        );
    }

    #[test]
    fn test_from_excel_date() {
        assert_eq!(
            from_excel_date(44516.42361),
            "2021-11-16T10:10:00+03:00".to_string()
        );
    }
}

use crate::simple_time::Moment;
use calamine::{open_workbook, DataType, Range, Reader, Xlsx, XlsxError};
use serde::{Deserialize, Serialize};
#[cfg(test)]
use std::env;
use std::{
    collections::HashMap,
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
const STATUS_NOT_GO: &str = "не идем";
const STATUS_ADMITTED: &str = "допущены";
const STATUS_APPLY: &str = "заявлены";
const STATUS_WIN: &str = "выиграли";
const STATUS_LOSS: &str = "не выиграли";
const STATUS_ESTIMATION: &str = "расчет";
const NAMED_RANGES_COUNT: usize = 20;
const RADIX: u32 = 36;

const TOTAL_COLUMNS: u32 = 16384;
// we don't expect that workbook
// has a number of active rows greater than that
const WORKBOOK_MAX_ROWS: usize = 7000;
// maximum days that we want to look in past
// searching for rows
const HOW_FAR_IN_PAST_DAYS: u64 = 16;

// 01.01.1970 in excel representation
const EXCEL_UNIX_EPOCH: u64 = 25569;

/// Alias result type for this module
type ExcelResult<T> = std::result::Result<T, WorkbookError>;

// WorkbookError is error type for this module
#[derive(Debug)]
pub enum WorkbookError {
    InvalidColumnNameError(String),
    XlsxError(XlsxError),
    JsonSerializeError(serde_json::Error),
}

impl From<serde_json::Error> for WorkbookError {
    fn from(error: serde_json::Error) -> Self {
        WorkbookError::JsonSerializeError(error)
    }
}

impl fmt::Display for WorkbookError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self {
            WorkbookError::InvalidColumnNameError(s) => write!(f, "invalid column name: {:?}", &s),
            WorkbookError::XlsxError(e) => write!(f, "{:?}", &e),
            WorkbookError::JsonSerializeError(e) => write!(f, "cannot serialize struct {:?}", e),
        }
    }
}

/// Prints to console active state records as json.
/// The 'activeness' of state is determined by [is_active state]
/// function
#[cfg(test)]
pub fn print_active_state_json(wb_path: &Path) -> ExcelResult<()> {
    let mut workbook = open_wb(wb_path)?;

    if let Some(purches) = active_purchases(&mut workbook) {
        println!("{}", purches.len());
        for p in purches {
            let j = serde_json::to_string_pretty(&p)?;
            println!("{}", &j);
        }
    }
    Ok(())
}

/// Returns a json string of active state records.
/// The 'activeness' of state is determined by [is_active state]
/// function
#[cfg(test)]
fn active_state_json(wb_path: &Path) -> ExcelResult<Option<String>> {
    let mut workbook = open_wb(wb_path)?;

    if let Some(purches) = active_purchases(&mut workbook) {
        Ok(Some(serde_json::to_string(&purches)?))
    } else {
        Ok(None)
    }
}

/// Converts vector of [Purchase] to json string
pub fn to_json(purches: &Vec<Purchase>) -> ExcelResult<String> {
    Ok(serde_json::to_string(purches)?)
}

/// Returns a result of comparing two sets of [Purchase]'s.
/// The first one is expected to be a [str] from some file
/// and considered as 'old'. The second one is expected to be
/// result of [active_state] function and considered as 'new'.
/// The goal is to get the records that has changed or simply new.
pub fn active_state_json_compared(
    old_purches: &str,
    new_purches: &Vec<Purchase>,
) -> ExcelResult<Option<String>> {
    // first we deserialize first set to a vector
    let old_purches: Vec<Purchase> = serde_json::from_str(old_purches)?;

    // second we build a hash set from that vector
    let mut old_purches_map: HashMap<String, Purchase> = HashMap::with_capacity(old_purches.len());
    for p in old_purches.into_iter() {
        old_purches_map.insert(p.registry_number.clone(), p); // this key is unique
    }

    let result = changed(&mut old_purches_map, new_purches);
    if result.len() == 0 {
        return Ok(None);
    }
    Ok(Some(serde_json::to_string(&result)?))
}

/// Compares two sets of data and returns resulting set
/// of records that has changed or records that is new
fn changed<'a>(
    old: &'a mut HashMap<String, Purchase>,
    new: &'a Vec<Purchase>,
) -> Vec<&'a Purchase> {
    let mut result: Vec<&Purchase> = Vec::new();

    for p in new {
        // if we have match on entries
        // we remove one from the first one
        match old.remove(&p.registry_number) {
            // than we compare if they are equal
            Some(v) if v == *p => continue, // if so we pass on next
            // if they are not equal or we didn't find
            // match than we push it to result
            _ => result.push(p),
        }
    }

    // if some records are left in first set
    // than this means that they are felt off
    // from active state and we need to include them
    // to the result with inactive state
    for (_, mut p) in old.into_iter() {
        p.status = STATUS_NOT_GO.to_string();
        result.push(p);
    }

    result
}

/// Prints to console active state records.
/// The 'activeness' of state is determined by [is_active state]
/// function
#[cfg(test)]
pub fn print_active_state(wb_path: &Path) -> ExcelResult<()> {
    let mut workbook = open_wb(wb_path)?;

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

/// Returns a vector of active state records.
/// The 'activeness' of state is determined by [is_active state]
/// function
pub fn active_state(wb_path: &Path) -> ExcelResult<Option<Vec<Purchase>>> {
    let mut workbook = open_wb(wb_path)?;

    Ok(active_purchases(&mut workbook))
}

/// Opens an excel workbook for further processing
fn open_wb(wb_path: &Path) -> ExcelResult<Xlsx<BufReader<File>>> {
    // let path = Path::new(path);
    Ok(open_workbook(wb_path).map_err(|x| WorkbookError::XlsxError(x))?)
}

/// Returns a vector of active state [Purchase]'s if any
fn active_purchases(workbook: &mut Xlsx<BufReader<File>>) -> Option<Vec<Purchase>> {
    // this function heavily relias on named ranges in workbook
    // and expected that they are equal to those that
    // processed by [is_expectable] function
    let named_ranges = named_ranges(&workbook);

    let named_cols = named_cols(&named_ranges);

    match workbook.worksheet_range(&named_ranges[0].sheet) {
        Some(Ok(range)) => {
            let purches = active_state_cells(range, named_cols);
            if !purches.is_empty() {
                Some(purches)
            } else {
                None
            }
        }
        _ => None,
    }
}

enum ColumnPosition {
    Left = 0,
    #[allow(dead_code)]
    Right = 1,
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

/// This type is represent a row in excel workbook
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

// [Purchase] equivalence logic
impl PartialEq for Purchase {
    fn eq(&self, other: &Self) -> bool {
        self.registry_number == other.registry_number
            && self.collecting_datetime == other.collecting_datetime
            && self.approval_datetime == other.approval_datetime
            && self.bidding_datetime == other.bidding_datetime
            && self.region == other.region
            && self.status == other.status
            && self.estimation as i64 == other.estimation as i64
            && self.our_participants == other.our_participants
            && self.winner == other.winner
            && self.winner_price as i64 == other.winner_price as i64
            && self.participants == other.participants
    }
}

struct NamedRange<'a> {
    name: &'a str,
    sheet: String, // sheet portion of 'List1!$A$:$A$' == 'List1'
    range: String, // range portion of 'List1!$A$:$A$' == '$A$:$A$'
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

    /// Gets a sheet portion of name
    fn parse_sheet(s: &str) -> String {
        s.split('!').nth(0).unwrap_or_default().to_owned()
    }

    /// Gets a range portion of name
    fn parse_range(s: &str) -> String {
        s.split('!').nth(1).unwrap_or_default().to_owned()
    }

    /// Returns column char name
    pub fn col_name(&self, cp: ColumnPosition) -> String {
        self.range
            .split(":")
            .nth(cp as usize)
            .unwrap_or_default()
            .replace("$", "")
    }

    /// Maps column char name to column serial number e.g. 'A' == 1
    pub fn col_num(&self, cp: ColumnPosition) -> ExcelResult<usize> {
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

/// Returns a vector of [NamedRange]'s build up from workbook defined names
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

/// Returns [NamedCols] struct that is represent result of
/// mapping [NamedRange] excel char column to serial number e.g. 'A' == 1
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

    // filter active rows and those that satisfy temporal criteria
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
        // if we get just time from excel cell i.e. cell value less than 1.0
        let bid_datetime = match r[cols.bidding_datetime].get_float().unwrap_or_default() {
            // than we take bidding_date
            d if d < 1.0 => r[cols.bidding_date].get_float().unwrap_or_default(),
            d => d,
        };
        let col_datetime = match r[cols.collecting_datetime].get_float().unwrap_or_default() {
            d if d < 1.0 => r[cols.collecting_date].get_float().unwrap_or_default(),
            d => d,
        };

        // get the values from the cells
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
        STATUS_ESTIMATION => true,
        _ => false,
    }
}

/// Converts an excel date which is just a float number
/// to formated time string as in RFS3339
fn from_excel_date(excel_date: f64) -> String {
    let days = excel_date - EXCEL_UNIX_EPOCH as f64;
    let seconds = seconds_from_days(days);
    let duration = Duration::new(seconds, 0);
    Moment::from_duration_since_epoch(duration).to_string()
}

/// Returns today's date in excel representation
/// which is a float number
fn today_in_excel_date() -> f64 {
    if let Ok(duration) = SystemTime::now().duration_since(SystemTime::UNIX_EPOCH) {
        let d_since_epoch = days_from_seconds(duration.as_secs());
        let today_excel = d_since_epoch + EXCEL_UNIX_EPOCH;
        today_excel as f64
    } else {
        0.0
    }
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
        let wb_path = env::var("REG_WORKBOOK_PATH").expect("$REG_WORKBOOK_PATH must be set");
        assert!(print_active_state(Path::new(&wb_path)).is_ok());
    }

    #[test]
    fn test_print_active_state_json() {
        let wb_path = env::var("REG_WORKBOOK_PATH").expect("$REG_WORKBOOK_PATH must be set");
        assert!(print_active_state_json(Path::new(&wb_path)).is_ok());
    }

    #[test]
    fn test_active_state_json() {
        let wb_path = env::var("REG_WORKBOOK_PATH").expect("$REG_WORKBOOK_PATH must be set");
        assert!(active_state_json(Path::new(&wb_path)).is_ok());
    }

    #[test]
    fn test_active_state() {
        let wb_path = env::var("REG_WORKBOOK_PATH").expect("$REG_WORKBOOK_PATH must be set");
        assert!(active_state(Path::new(&wb_path)).is_ok());
    }

    #[test]
    fn test_open_wb() {
        let wb_path = env::var("REG_WORKBOOK_PATH").expect("$REG_WORKBOOK_PATH must be set");
        assert!(open_wb(Path::new(&wb_path)).is_ok());
    }

    #[test]
    fn test_from_excel_date() {
        assert_eq!(
            from_excel_date(44516.42361),
            "2021-11-16T10:10:00+00:00".to_string()
        );
        assert_eq!(
            from_excel_date(44621.00),
            "2022-03-01T00:00:00+00:00".to_string()
        );
        assert_eq!(
            from_excel_date(43890.4166666667),
            "2020-02-29T10:00:00+00:00".to_string()
        );
        assert_eq!(
            from_excel_date(44713.00),
            "2022-06-01T00:00:00+00:00".to_string()
        );
    }

    #[test]
    fn test_is_active_state() {
        assert!(is_active_state(STATUS_GO));
        assert!(is_active_state(STATUS_ADMITTED));
        assert!(is_active_state(STATUS_APPLY));
        assert!(is_active_state(STATUS_WIN));
        assert!(is_active_state(STATUS_LOSS));
        assert!(is_active_state(STATUS_ESTIMATION));
        assert!(!is_active_state("invalid status"))
    }

    #[test]
    fn test_is_expectable() {
        assert!(is_expectable(DATE_COLLECTING_BIDS));
        assert!(is_expectable(TIME_COLLECTING_BIDS));
        assert!(is_expectable(DATE_APPROVAL));
        assert!(is_expectable(DATE_BIDDING));
        assert!(is_expectable(TIME_BIDDING));
        assert!(is_expectable(MAX_PRICE));
        assert!(is_expectable(NUMBER));
        assert!(is_expectable(WINNER));
        assert!(is_expectable(STATUS));
        assert!(is_expectable(WINNER_PRICE));
        assert!(is_expectable(PARTICIPANTS));
        assert!(is_expectable(PURCHASE_TYPE));
        assert!(is_expectable(PURCHASE_SUBJECT));
        assert!(is_expectable(PURCHASE_SUBJECT_ABBR));
        assert!(is_expectable(OUR_PARTICIPANTS));
        assert!(is_expectable(ETP));
        assert!(is_expectable(APPLICATION_GUARANTEE));
        assert!(is_expectable(CONTRACT_GUARANTEE));
        assert!(is_expectable(CUSTOMER_TYPE));
        assert!(is_expectable(ESTIMATION));
        assert!(is_expectable(REGION));
        assert!(!is_expectable("invalid name"))
    }
}

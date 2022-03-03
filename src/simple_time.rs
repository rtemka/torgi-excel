use std::time::{Duration, SystemTime};

const JANUARY: (u32, &str, u32) = (1, "January", 31);
const FEBRUARY: (u32, &str, u32) = (2, "February", 28);
const MARCH: (u32, &str, u32) = (3, "March", 31);
const APRIL: (u32, &str, u32) = (4, "April", 30);
const MAY: (u32, &str, u32) = (5, "May", 31);
const JUNE: (u32, &str, u32) = (6, "June", 30);
const JULY: (u32, &str, u32) = (7, "July", 31);
const AUGUST: (u32, &str, u32) = (8, "August", 31);
const SEPTEMBER: (u32, &str, u32) = (9, "September", 30);
const OCTOBER: (u32, &str, u32) = (10, "October", 31);
const NOVEMBER: (u32, &str, u32) = (11, "November", 30);
const DECEMBER: (u32, &str, u32) = (12, "December", 31);

// /// Enum of months tuples in format: (number, name, days in)
// enum Month<'a> {
//     Ordinary((u32, &'a str, u32)),
//     February((u32, &'a str, u32)),
// }

/// Returns option of month tuple. If number of month is not in a range of 1..12 then returns None
pub fn month<'a>(month: usize) -> Option<(u32, &'a str, u32)> {
    match month {
        1 => Some(JANUARY),
        2 => Some(FEBRUARY),
        3 => Some(MARCH),
        4 => Some(APRIL),
        5 => Some(MAY),
        6 => Some(JUNE),
        7 => Some(JULY),
        8 => Some(AUGUST),
        9 => Some(SEPTEMBER),
        10 => Some(OCTOBER),
        11 => Some(NOVEMBER),
        12 => Some(DECEMBER),
        _ => None,
    }
}

/// Struct built from SystemTime (now or from passed in). It is UTC time, no local time.
#[derive(Debug)]
pub struct Moment {
    pub year: u64,
    pub month: u64,
    pub day: u64,
    pub hours: u64,
    pub minutes: u64,
    pub seconds: u64,
    pub is_leap_year: bool,
}

impl Moment {
    /// Returns option of Moment from now
    pub fn new() -> Option<Self> {
        SystemTime::now()
            .duration_since(SystemTime::UNIX_EPOCH)
            .ok()
            .and_then(|d| Some(Self::from_duration_since_epoch(d)))
    }

    /// Returns option of Moment. None if timestamp is before UNIX epoch
    pub fn from_sys_time(time: SystemTime) -> Option<Self> {
        time.duration_since(SystemTime::UNIX_EPOCH)
            .ok()
            .and_then(|d| Some(Self::from_duration_since_epoch(d)))
    }

    pub const fn from_duration_since_epoch(dse: Duration) -> Self {
        let duration_in_secs = dse.as_secs();

        let year = year(duration_in_secs);
        let extra_days = extra_days(year);
        let is_leap_year = is_leap_year(year);

        let d_since_epoch = days_since_epoch(duration_in_secs);
        let d_in_current_year = days_in_current_year(d_since_epoch, extra_days) + 1;

        let sec_in_day = seconds_in_day(duration_in_secs, d_since_epoch);
        let hrs = hours(sec_in_day);
        let mins = minutes(sec_in_day, hrs);
        let secs = seconds(sec_in_day, hrs, mins);

        let (day, month) = day_and_month(d_in_current_year, is_leap_year);
        Self {
            year,
            month,
            day,
            hours: hrs,
            minutes: mins,
            seconds: secs,
            is_leap_year,
        }
    }
}

/// Unix timestamp / hours in a year to get years from 1970 to timestamp
const fn year(timestamp_secs: u64) -> u64 {
    //31436000 hours in a year
    (timestamp_secs / 31536000) + 1970
}

/// Determine number of extra days from leap years since 1970
const fn extra_days(year: u64) -> u64 {
    (year - 1969) / 4
}

/// Is current year a leap year?
const fn is_leap_year(year: u64) -> bool {
    (year - 1969) % 4 == 3
}

/// Determine the number of days since the epoch
const fn days_since_epoch(timestamp_secs: u64) -> u64 {
    timestamp_secs / 86400
}

/// Modulo the (days since epoch - extra days from leap years) by the number of days in a year to find the days passed in the current year
const fn days_in_current_year(days_since_epoch: u64, extra_days: u64) -> u64 {
    (days_since_epoch - extra_days) % 365
}

/// Find the number of seconds in the current day.
const fn seconds_in_day(timestamp_secs: u64, days_since_epoch: u64) -> u64 {
    timestamp_secs - (days_since_epoch * 86400)
}

/// Find the number of hours in the current day.
const fn hours(seconds_in_day: u64) -> u64 {
    seconds_in_day / 3600
}

/// Find the number of minutes in the current hour.
const fn minutes(seconds_in_day: u64, hours: u64) -> u64 {
    (seconds_in_day - (hours * 3600)) / 60
}

/// Find the number of seconds in the current minute.
const fn seconds(seconds_in_day: u64, hours: u64, minutes: u64) -> u64 {
    (seconds_in_day - (hours * 3600)) - (minutes * 60)
}

/// Returns the month and day of that month
const fn day_and_month(days_in_current_year: u64, is_leap_year: bool) -> (u64, u64) {
    let add_day = is_leap_year as u64;
    match days_in_current_year {
        d if d <= 31 => (d, JANUARY.0 as u64),
        d if d <= (60 + add_day) => (d - 31, FEBRUARY.0 as u64),
        d if d <= (91 + add_day) => (d - (59 + add_day), MARCH.0 as u64),
        d if d <= (121 + add_day) => (d - (90 + add_day), APRIL.0 as u64),
        d if d <= (152 + add_day) => (d - (120 + add_day), MAY.0 as u64),
        d if d <= (182 + add_day) => (d - (151 + add_day), JUNE.0 as u64),
        d if d <= (213 + add_day) => (d - (181 + add_day), JULY.0 as u64),
        d if d <= (244 + add_day) => (d - (212 + add_day), AUGUST.0 as u64),
        d if d <= (274 + add_day) => (d - (243 + add_day), SEPTEMBER.0 as u64),
        d if d <= (305 + add_day) => (d - (273 + add_day), OCTOBER.0 as u64),
        d if d <= (335 + add_day) => (d - (304 + add_day), NOVEMBER.0 as u64),
        d if d <= (365 + add_day) => (d - (334 + add_day), DECEMBER.0 as u64),
        _ => (0, 0),
    }
}

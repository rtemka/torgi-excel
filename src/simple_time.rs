use std::fmt;
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

/// Returns option of month tuple. If number of month is not in a range of 1..12 then returns None
#[cfg(test)]
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

/// Struct built from SystemTime (now or from passed in). It is GMT time, no local offsets.
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

        let y = year(duration_in_secs);
        let extra_days = extra_days(y);
        let is_leap_year = is_leap_year(y);

        let d_since_epoch = days_since_epoch(duration_in_secs);
        let d_in_current_year = days_in_current_year(y, d_since_epoch, extra_days);

        let sec_in_day = seconds_in_day(duration_in_secs, d_since_epoch);
        let hrs = hours(sec_in_day);
        let mins = minutes(sec_in_day, hrs);
        let secs = seconds(sec_in_day, hrs, mins);

        let (day, month) = day_and_month(d_in_current_year, is_leap_year);
        Self {
            year: y,
            month,
            day,
            hours: hrs,
            minutes: mins,
            seconds: secs,
            is_leap_year,
        }
    }

    /// Adds leading zeroes to the unit of time
    /// for the proper displaying
    fn add_leading_zero(x: u64) -> String {
        if x < 10 {
            return format!("0{}", x);
        }
        x.to_string()
    }
}

impl fmt::Display for Moment {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        // we want to build rfc3339 format string
        // like "2006-01-02T15:04:05Z00:00"
        write!(
            f,
            "{}-{}-{}T{}:{}:{}+00:00",
            self.year,
            Moment::add_leading_zero(self.month),
            Moment::add_leading_zero(self.day),
            Moment::add_leading_zero(self.hours),
            Moment::add_leading_zero(self.minutes),
            Moment::add_leading_zero(self.seconds),
        )
    }
}

/// Unix timestamp / hours in a year to get years from 1970 to timestamp
const fn year(timestamp_secs: u64) -> u64 {
    //approximately 31536000 hours in a year
    let ext_d = extra_days(timestamp_secs / 31536000 + 1970);
    (timestamp_secs - ext_d * 86400) / 31536000 + 1970
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

/// Subtract days from leap years and days from average years from days since epoch
/// to find the days passed in the current year
const fn days_in_current_year(year: u64, days_since_epoch: u64, leap_years: u64) -> u64 {
    // days from leap year plus days from average years
    let days_till_this_year = ((year - 1970 - leap_years) * 365) + (leap_years * 366);
    days_since_epoch - days_till_this_year + 1
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
        d if d <= (59 + add_day) => (d - 31, FEBRUARY.0 as u64),
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_add_leading_zero() {
        assert_eq!(Moment::add_leading_zero(9), "09".to_string());
        assert_eq!(Moment::add_leading_zero(11), "11".to_string());
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
        assert_eq!(m.to_string(), "2021-11-10T12:01:44+00:00".to_string());
    }

    #[test]
    fn test_month() {
        assert_eq!(Some(JANUARY), month(1));
        assert_eq!(Some(FEBRUARY), month(2));
        assert_eq!(Some(MARCH), month(3));
        assert_eq!(Some(APRIL), month(4));
        assert_eq!(Some(MAY), month(5));
        assert_eq!(Some(JUNE), month(6));
        assert_eq!(Some(JULY), month(7));
        assert_eq!(Some(AUGUST), month(8));
        assert_eq!(Some(SEPTEMBER), month(9));
        assert_eq!(Some(OCTOBER), month(10));
        assert_eq!(Some(NOVEMBER), month(11));
        assert_eq!(Some(DECEMBER), month(12));
        assert_eq!(None, month(99));
        assert_eq!(None, month(13));
    }

    #[test]
    fn test_day_and_month() {
        assert_eq!((31, 1), day_and_month(31, false));
        assert_eq!((29, 2), day_and_month(60, true));
        assert_eq!((1, 3), day_and_month(60, false));
        assert_eq!((1, 3), day_and_month(61, true));
        assert_eq!((28, 2), day_and_month(59, false));
        assert_eq!((31, 12), day_and_month(365, false));
        assert_eq!((31, 12), day_and_month(366, true));
    }

    #[test]
    fn test_year() {
        assert_eq!(2022, year(1646815260));
        assert_eq!(2021, year(1640908740));
        assert_eq!(2020, year(1583743260));
        assert_eq!(2009, year(1262303999));
        assert_eq!(2010, year(1262304000));
    }

    #[test]
    fn test_from_duration_since_epoch() {
        let d = Duration::from_secs(1262303999);
        let m = Moment::from_duration_since_epoch(d);
        assert_eq!(
            (2009, 12, 31, 23, 59, 59),
            (m.year, m.month, m.day, m.hours, m.minutes, m.seconds)
        );

        let d = Duration::from_secs(1262304000);
        let m = Moment::from_duration_since_epoch(d);
        assert_eq!(
            (2010, 1, 1, 0, 0, 0),
            (m.year, m.month, m.day, m.hours, m.minutes, m.seconds)
        );

        let d = Duration::from_secs(1583020799);
        let m = Moment::from_duration_since_epoch(d);
        assert_eq!(
            (2020, 2, 29, 23, 59, 59),
            (m.year, m.month, m.day, m.hours, m.minutes, m.seconds)
        );

        let d = Duration::from_secs(1583020800);
        let m = Moment::from_duration_since_epoch(d);
        assert_eq!(
            (2020, 3, 1, 0, 0, 0),
            (m.year, m.month, m.day, m.hours, m.minutes, m.seconds)
        );
    }

    #[test]
    fn test_extra_days() {
        assert_eq!(14, extra_days(2025));
        assert_eq!(13, extra_days(2021));
        assert_eq!(12, extra_days(2020));
        assert_eq!(11, extra_days(2016));
        assert_eq!(10, extra_days(2012));
        assert_eq!(9, extra_days(2008));
        assert_eq!(8, extra_days(2004));
        assert_eq!(0, extra_days(1971));
        assert_eq!(1, extra_days(1973));
        assert_eq!(2, extra_days(1977));
    }

    #[test]
    fn test_is_leap_year() {
        assert_eq!(true, is_leap_year(2020));
        assert_eq!(false, is_leap_year(2021));
        assert_eq!(false, is_leap_year(2013));
        assert_eq!(true, is_leap_year(2016));
        assert_eq!(true, is_leap_year(2012));
        assert_eq!(true, is_leap_year(2008));
        assert_eq!(true, is_leap_year(2004));
        assert_eq!(true, is_leap_year(1980));
        assert_eq!(true, is_leap_year(1972));
        assert_eq!(true, is_leap_year(1976));
    }

    #[test]
    fn test_days_in_current_year() {
        let y = year(1634328000);
        let ext_days = extra_days(y);
        let ds = days_since_epoch(1634328000);

        assert_eq!(288, days_in_current_year(y, ds, ext_days));

        let y = year(1640995199);
        let ext_days = extra_days(y);
        let ds = days_since_epoch(1640995199);

        assert_eq!(365, days_in_current_year(y, ds, ext_days));

        let y = year(1609372800);
        let ext_days = extra_days(y);
        let ds = days_since_epoch(1609372800);

        assert_eq!(366, days_in_current_year(y, ds, ext_days));

        let y = year(1577836800);
        let ext_days = extra_days(y);
        let ds = days_since_epoch(1577836800);

        assert_eq!(1, days_in_current_year(y, ds, ext_days));

        let y = year(1582934400);
        let ext_days = extra_days(y);
        let ds = days_since_epoch(1582934400);

        assert_eq!(60, days_in_current_year(y, ds, ext_days));

        let y = year(1609286400);
        let ext_days = extra_days(y);
        let ds = days_since_epoch(1609286400);

        assert_eq!(365, days_in_current_year(y, ds, ext_days));
    }

    #[test]
    fn test_days_since_epoch() {
        assert_eq!(18992, days_since_epoch(1640908800));
        assert_eq!(0, days_since_epoch(0));
    }
}

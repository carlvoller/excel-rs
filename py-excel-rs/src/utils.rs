use chrono::{NaiveDate, NaiveDateTime, NaiveTime};

const NANOSECONDS_IN_A_DAY: f64 = 1_000_000_000.0 * 60.0 * 60.0 * 24.0;

// const EXCEL1900_EPOCH: &str = "1899-12-30 00:00:00";
// const EXCEL1904_EPOCH: &str = "1904-01-01 00:00:00";
// const EXCEL_MIN_TIME1900_EPOCH: &str = "1899-12-31 00:00:00";
// const EXCEL_BUGGY_DATE_START: &str = "1900-03-01 00:00:00";
// const DATE_TIME_FORMAT: &str = "%Y-%m-%d %H:%M:%S";

pub fn chrono_to_xlsx_date(date: NaiveDateTime) -> f64 {
    let start_date_only = NaiveDate::from_ymd_opt(1900, 1, 1).expect("Excel Epoch");
    let start_date = NaiveDateTime::new(
        start_date_only,
        NaiveTime::from_hms_opt(0, 0, 0).expect("Time Epoch"),
    );
    let delta = (date - start_date).num_nanoseconds().unwrap();

    let true_delta: f64 = delta as f64 / NANOSECONDS_IN_A_DAY;
    return true_delta;
}

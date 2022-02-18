use log::{error, info};
use reqwest::{blocking::Client, header::CONTENT_TYPE};
use std::{
    fmt, fs,
    fs::File,
    io::{Read, Write},
    path::Path,
    thread,
    time::{self, SystemTime},
};

use crate::excel;

/// Daemon sleep interval in seconds
const TIME_TO_SLEEP: u64 = 10 * 60;

const TEMP_FILE_PATH: &str = "temp.json";

/// Returns last modification time of a file or [std::io::Error]
fn last_modified_time(path: &Path) -> std::io::Result<SystemTime> {
    fs::metadata(path).and_then(|m| m.modified())
}

fn file_content(path: &Path) -> std::io::Result<String> {
    let mut file = File::open(&path)?;
    let mut s = String::new();
    file.read_to_string(&mut s)?;
    Ok(s)
}

fn write_to_file(path: &Path, s: String) -> std::io::Result<()> {
    let mut file = File::create(&path)?;
    file.write_all(s.as_bytes())?;
    Ok(())
}

fn send(client: &Client, json: String) -> Result<(), reqwest::Error> {
    let res = client
        .post(crate::APP_URL)
        .header(CONTENT_TYPE, "application/json")
        .body(json)
        .send();

    match res {
        Ok(r) => {
            info!("update is sent; response status: {}", r.status());
            Ok(())
        }
        Err(e) => {
            error!("error while sending update: {:?}", &e);
            Err(e)
        }
    }
}

pub fn watch(file_path: &str) -> Result<(), DaemonError> {
    let path = Path::new(file_path);
    let temp_path = Path::new(TEMP_FILE_PATH);

    let sleep_time = time::Duration::from_secs(TIME_TO_SLEEP);

    let mut last_mod_time = last_modified_time(&path)?;

    info!("last modification time is set to: {:?}", last_mod_time);

    let client = Client::new();

    loop {
        thread::sleep(sleep_time);

        let time_checked = last_modified_time(&path)?;

        if time_checked.eq(&last_mod_time) {
            continue;
        }

        info!("file change detected");

        let json = if temp_path.exists() {
            let content = file_content(temp_path)?;
            match excel::active_state_json_compared(&content) {
                Ok(Some(s)) => s,
                Ok(None) => continue,
                Err(e) => return Err(e.into()),
            }
        } else {
            match excel::active_state_json() {
                Ok(Some(s)) => s,
                Ok(None) => continue,
                Err(e) => return Err(e.into()),
            }
        };

        if let Err(_) = send(&client, json.clone()) {
            continue;
        }

        write_to_file(temp_path, json)?;
        last_mod_time = time_checked;
    }
}

#[derive(Debug)]
pub enum DaemonError {
    IoError(std::io::Error),
    WorkBookError(excel::WorkbookError),
}

impl From<std::io::Error> for DaemonError {
    fn from(error: std::io::Error) -> Self {
        DaemonError::IoError(error)
    }
}

impl From<excel::WorkbookError> for DaemonError {
    fn from(error: excel::WorkbookError) -> Self {
        DaemonError::WorkBookError(error)
    }
}

impl fmt::Display for DaemonError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self {
            DaemonError::WorkBookError(e) => write!(f, "{:?}", &e),
            DaemonError::IoError(e) => write!(f, "{:?}", &e),
        }
    }
}

use crate::{excel, simple_time};
use ctrlc;
use log::{error, info};
use reqwest::{blocking::Client, header::CONTENT_TYPE};
use std::{
    fmt, fs,
    fs::File,
    io::{Read, Write},
    path::Path,
    sync::{Arc, Mutex},
    thread,
    time::{self, SystemTime},
};

/// Daemon sleep interval in seconds
const TIME_TO_SLEEP: u64 = 30;

const TEMP_FILE_PATH: &str = "temp.json";

/// Returns last modification time of a file or [std::io::Error]
fn last_modified_time(path: &Path) -> std::io::Result<SystemTime> {
    fs::metadata(path).and_then(|m| m.modified())
}

/// Reads content of the file to string
fn file_content(path: &Path) -> std::io::Result<String> {
    let mut file = File::open(&path)?;
    let mut s = String::new();
    file.read_to_string(&mut s)?;
    Ok(s)
}

/// Writes string to file
fn write_to_file(path: &Path, s: String) -> std::io::Result<()> {
    let mut file = File::create(&path)?;
    file.write_all(s.as_bytes())?;
    Ok(())
}

/// Removes program temporary file
fn rm_file(temp_path: &Path) -> std::io::Result<()> {
    if temp_path.exists() {
        fs::remove_file(temp_path)?
    }
    Ok(())
}

/// Sends json post request to remote app url
fn send(client: &Client, url: &str, json: String) -> Result<(), reqwest::Error> {
    let res = client
        .post(url)
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

fn print_time(sys_time: SystemTime) {
    match simple_time::Moment::from_sys_time(sys_time) {
        Some(m) => info!("last modification time is: {}", m.to_string()),
        None => info!("couldn't parse time from system time"),
    };
}

/// Ctrl+C signal handling logic
fn interrupt_handler(sig_flag: Arc<Mutex<bool>>) -> Result<(), DaemonError> {
    let handler = move || {
        info!("received Ctrl+C, exiting...");
        *sig_flag.lock().unwrap() = true;
    };
    ctrlc::set_handler(handler).map_err(|_| DaemonError::SetSignalError)?;
    Ok(())
}

/// Checks a file for changes every time that is specified by [TIME_TO_SLEEP].
/// This daemon has it's own litlle presistent store which is a file with
/// result of previous checking. So if the change of the file is detected,
/// than it either compare previous result with the new one or just take new one
/// and send it to the remote app url (which is where database resides)
pub fn watch(file_path: &str, to_send_url: &str) -> Result<(), DaemonError> {
    let path = Path::new(file_path);
    let temp_path = Path::new(TEMP_FILE_PATH); // path of the storage file

    let sleep_time = time::Duration::from_secs(TIME_TO_SLEEP);

    let mut last_mod_time = last_modified_time(&path)?;

    let client = Client::new();

    // Ctrl+C handling
    let interrupt_sig_handler = Arc::new(Mutex::new(false));
    let interrupt_sig_main = interrupt_sig_handler.clone();
    interrupt_handler(interrupt_sig_handler)?;

    info!("start watching to '{}'", &file_path);

    while *interrupt_sig_main.lock().unwrap() != true {
        thread::sleep(sleep_time);

        let time_checked = last_modified_time(&path)?;

        if time_checked.eq(&last_mod_time) {
            continue;
        }

        info!("file change detected");

        // we get active state records from the file
        let new_snapshot = match excel::active_state(&path) {
            Ok(Some(s)) => s,
            Ok(None) => continue,
            Err(e) => return Err(e.into()),
        };

        // if we have a temp file with previous records
        // than we compare old with new and gets result set
        let json = if temp_path.exists() {
            let old_snapshot = file_content(temp_path)?;
            match excel::active_state_json_compared(&old_snapshot, &new_snapshot) {
                Ok(Some(s)) => s,
                Ok(None) => {
                    info!("no changes in records");
                    last_mod_time = time_checked;
                    print_time(last_mod_time);
                    continue;
                }
                Err(e) => return Err(e.into()),
            }
        } else {
            excel::to_json(&new_snapshot)? // or we just take new records
        };

        // if we have an error from remote database
        // than error is logged by send function
        if let Err(_) = send(&client, to_send_url, json) {
            continue;
        }

        let new_snapshot = excel::to_json(&new_snapshot)?;

        write_to_file(temp_path, new_snapshot)?;
        last_mod_time = time_checked;
        print_time(last_mod_time);
    }

    info!("removing any temp files...");
    rm_file(temp_path)?; // remove temp files

    Ok(())
}

/// DaemonError is the wrapper around
/// either [std::io::Error] or [excel::WorkbookError]
#[derive(Debug)]
pub enum DaemonError {
    SetSignalError,
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
            DaemonError::SetSignalError => write!(f, "error setting Ctrl-C handler"),
        }
    }
}

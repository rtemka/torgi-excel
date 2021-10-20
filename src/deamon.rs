use env_logger;
use log::{error, info};
use reqwest::{blocking::Client, header::CONTENT_TYPE};
use std::{env, fs, path::Path, thread, time};

use crate::excel;

const TIME_TO_SLEEP: u64 = 20;

fn enable_logger() {
    env::set_var("RUST_LOG", "info");
    env_logger::init();
}

pub fn send_when_modify(file_path: &str) {
    enable_logger();
    let path = Path::new(file_path);

    let sleep_time = time::Duration::from_secs(TIME_TO_SLEEP);

    let mut last_mod_time = match fs::metadata(path) {
        Ok(m) => match m.modified() {
            Ok(time) => time,
            Err(e) => {
                error!("{:?}", &e);
                return;
            }
        },
        Err(e) => {
            error!("{:?}", &e);
            return;
        }
    };
    info!("last modification time is set to: {:?}", last_mod_time);

    let client = Client::new();

    loop {
        thread::sleep(sleep_time);

        let time = match fs::metadata(path) {
            Ok(m) => match m.modified() {
                Ok(time) => time,
                Err(e) => {
                    error!("{:?}", &e);
                    return;
                }
            },
            Err(e) => {
                error!("{:?}", &e);
                return;
            }
        };

        if time.ne(&last_mod_time) {
            info!("file change detected");

            last_mod_time = time;

            let active_state_records = excel::get_active_state_json().unwrap();
            let res = client
                .post(crate::APP_URL)
                .header(CONTENT_TYPE, "application/json")
                .body(active_state_records)
                .send();
            match res {
                Ok(r) => info!("update is sent; response status: {}", r.status()),
                Err(e) => error!("error while sending update: {:?}", &e),
            }
        }
    }
}

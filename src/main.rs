mod daemon;
mod excel;
mod simple_time;
use env_logger;
use log::{error, info};
use std::env;

const WORKBOOK_PATH: &str = "//rsphnas/Inbox/упр.мод/Форматы/Форматы/Форматы отд. торгов/Форматы отд. торгов/Реестр 2022.xlsx";
const APP_URL: &str = "https://torgi-contracts-bot.herokuapp.com/KMZ4aV0pffnvepuQY3YsGIYghtsy1Thq";

/// Logger initialization
fn enable_logger() {
    env::set_var("RUST_LOG", "info");
    env_logger::init();
}

fn main() {
    enable_logger();
    match daemon::watch(WORKBOOK_PATH) {
        Ok(()) => info!("done"),
        Err(e) => error!("{:?}", &e),
    };
    // let p = excel::active_state_json().unwrap().unwrap();
    // println!("{:?}", p)
}

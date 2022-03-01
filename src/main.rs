mod daemon;
mod excel;
mod simple_time;
use env_logger;
use log::{error, info};
use std::env;

/// Logger initialization
fn enable_logger() {
    env::set_var("RUST_LOG", "info");
    env_logger::init();
}

fn main() {
    enable_logger();

    let wb_path = env::var("REG_WORKBOOK_PATH").expect("$REG_WORKBOOK_PATH must be set");
    let app_url = env::var("TGBOT_APP_URL").expect("$TGBOT_APP_URL must be set");

    match daemon::watch(&wb_path, &app_url) {
        Ok(()) => info!("done"),
        Err(e) => error!("{:?}", &e),
    };
}

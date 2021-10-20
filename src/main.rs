mod deamon;
mod excel;

const WORKBOOK_PATH: &str = "//rsphnas/Inbox/упр.мод/Форматы/Форматы/Форматы отд. торгов/Форматы отд. торгов/Реестр 2021.xlsx";
const APP_URL: &str = "https://torgi-contracts-bot.herokuapp.com/KMZ4aV0pffnvepuQY3YsGIYghtsy1Thq";

fn main() {
    deamon::send_when_modify(WORKBOOK_PATH);
}

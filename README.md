### **Программа для отслеживания excel-файла**

- Работает в связке с [телеграм-ботом](https://github.com/rtemka/torgi-contracts-bot)

- Программа следит за изменениями excel-файла
- Путь к файлу программа получает из переменной окружения
```bash
export REG_WORKBOOK_PATH="path/to/excel/file"
```
- Если в файле произошли изменения, то эти изменения сериализуются в формат JSON и отсылаются на API бота
```bash
export TGBOT_APP_URL="https://[app-name].herokuapp/[db-update-token]"
```
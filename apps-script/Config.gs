/**
 * Константы сервиса «Активности».
 * ID таблицы и папки с оборотами можно переопределить в свойствах скрипта
 * (SPREADSHEET_ID, DATA_FOLDER_ID), не меняя код.
 */
var APP = {
  SPREADSHEET_ID: '1_NVSx7JIXOe9bwoEjsOn93T1ZbQsrq2ijJi0QBVH9sY',
  DATA_FOLDER_ID: '1tXCRf_EUS8fI-Feonv6vWmfX2gbMBCog',
  ADMIN_FIO: 'Киндсфатер Владимир Владимирович',
  // Адрес веб-приложения (clasp deploy). Свойство скрипта APP_URL или настройка в админке имеют приоритет.
  WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbzyzTvd1_YNbMA7BnFK_pxpnBHJq6tP_6cceSkogJd8BnlgBLnO9OwBEBozrt9n9GjG/exec',
  TZ: 'Europe/Kaliningrad',
  MIN_COMMENT: 10,
  SHEET: {
    ACT: 'Активности',
    USERS: 'Пользователи',
    DICT: 'Справочники',
    TURN: 'Обороты',
    COMP: 'Компании',
    LOG: 'Журнал'
  }
};

// Колонки листа «Активности». A–L — исходная структура файла, M–S — служебные.
var COL = {
  FIO: 1, NO: 2, INN: 3, SHORT: 4, FULL: 5, CHANNEL: 6, POTENTIAL: 7, STATUS: 8,
  DATE: 9, TYPE: 10, NEXT: 11, COMMENT: 12, DEAL_SUM: 13,
  ID: 14, POT_SRC: 15, AUTHOR: 16, CREATED: 17, UPDATED: 18, DELETED: 19
};
var NCOL = 19;
var ACT_HEADERS = [
  'ФИО МПП/МПС/МРГС/НУ', 'ФИО НО', 'ИНН', 'Краткое название ЮЛ', 'Полное название ЮЛ',
  'Канал поступления', 'Потенциал оборота в мес', 'Статус', 'Дата активности', 'Вид активности',
  'Дата следующей активности', 'Комментарий', 'Сумма сделки',
  'ID', 'Источник потенциала', 'Автор', 'Создано', 'Изменено', 'Удалено'
];

var ROLES = { ADMIN: 'Администратор', HEAD: 'Начальник отдела', EMPLOYEE: 'Сотрудник' };
var EMP_TYPES = ['МПП', 'МПС', 'МРГК', 'НУ'];

var UCOL = { ID: 1, FIO: 2, TYPE: 3, NO: 4, ROLE: 5, REPORTS: 6, EXTRA: 7, ACTIVE: 8, TOKEN: 9, LINK: 10 };
var USER_HEADERS = [
  'ID', 'ФИО', 'Тип сотрудника', 'ФИО НО', 'Роль', 'Отчёты', 'Доп. доступ к ФИО (через ;)',
  'Активен', 'Токен', 'Персональная ссылка'
];

// Типы статусов: open — клиент в работе, deal — сделка (нужна сумма), lost — закрыт без сделки.
var KIND_LABEL = { open: 'Открыт', deal: 'Сделка', lost: 'Закрыт' };
var DEFAULT_DICTS = {
  channels: ['Кампания продаж', 'Пиратка', 'Полевой обход', 'Личное привлечение', 'Другое'],
  statuses: [
    { name: 'В работе', kind: 'open' },
    { name: 'Недозвон', kind: 'open' },
    { name: 'Отказ', kind: 'lost' },
    { name: 'Блокировка 115-ФЗ', kind: 'lost' },
    { name: 'Договор подписан', kind: 'deal' },
    { name: 'Договор запущен', kind: 'deal' },
    { name: 'Выполнено расширение', kind: 'deal' }
  ],
  types: ['Звонок', 'Встреча', 'Сообщение']
};
var DICT_HEADERS = ['Канал поступления', 'Статус', 'Тип статуса (Открыт/Сделка/Закрыт)', 'Вид активности'];

var TURN_HEADERS = [
  'ИНН', 'Оборот конкурентов, мес (PS)', 'Банки конкурентов', 'Доля СБП', 'Доля e-com',
  'Оборот Сбер, мес', 'Комиссионный доход Сбер', 'Файл PS', 'Файл Сбер', 'Обновлено'
];
var COMP_HEADERS = ['ИНН', 'Краткое название', 'Полное название', 'ОГРН', 'Регион', 'Статус', 'Источник', 'Обновлено'];
var LOG_HEADERS = ['Время', 'Пользователь', 'Действие', 'ID', 'Детали'];

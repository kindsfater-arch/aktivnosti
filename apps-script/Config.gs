/**
 * Константы сервиса «Активности».
 * ID таблицы и папки с оборотами можно переопределить в свойствах скрипта
 * (SPREADSHEET_ID, DATA_FOLDER_ID), не меняя код.
 */
var APP = {
  SPREADSHEET_ID: '1_NVSx7JIXOe9bwoEjsOn93T1ZbQsrq2ijJi0QBVH9sY',
  DATA_FOLDER_ID: '1tXCRf_EUS8fI-Feonv6vWmfX2gbMBCog',
  ADMIN_FIO: 'Киндсфатер Владимир Владимирович',
  // Адрес интерфейса на GitHub Pages — на него ведут персональные ссылки.
  PUBLIC_URL: 'https://kindsfater-arch.github.io/aktivnosti/',
  // Адрес веб-приложения Apps Script (API). Свойство скрипта APP_URL или настройка в админке имеют приоритет.
  WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbzyzTvd1_YNbMA7BnFK_pxpnBHJq6tP_6cceSkogJd8BnlgBLnO9OwBEBozrt9n9GjG/exec',
  TZ: 'Europe/Kaliningrad',
  MIN_COMMENT: 10,
  SHEET: {
    ACT: 'Активности',
    USERS: 'Пользователи',
    DICT: 'Справочники',
    TURN: 'Обороты',
    COMP: 'Компании',
    PLANS: 'Планы',
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
// active: false — статус остаётся в отчётах и старых записях, но выбрать его в новой или изменённой активности нельзя.
var KIND_LABEL = { open: 'Открыт', deal: 'Сделка', lost: 'Закрыт' };
var DEFAULT_DICTS = {
  channels: ['Кампания продаж', 'Пиратка', 'Полевой обход', 'Личное привлечение', 'Другое'],
  statuses: [
    { name: 'Недозвон', kind: 'open', active: true },
    { name: 'Отказ', kind: 'lost', active: true },
    { name: 'Договор подписан', kind: 'deal', active: true },
    { name: 'Выполнено расширение', kind: 'deal', active: true },
    { name: 'Нет контактов', kind: 'open', active: true },
    { name: 'Назначено', kind: 'open', active: true },
    { name: 'Получено согласие', kind: 'open', active: true },
    { name: 'Клиент думает', kind: 'open', active: true }
  ],
  types: ['Звонок', 'Встреча', 'Сообщение']
};
// Увеличьте, чтобы статусы из DEFAULT_DICTS (тип и доступность) один раз заново записались в справочник.
var DICT_VERSION = 2;
var DICT_HEADERS = ['Канал поступления', 'Статус', 'Тип статуса (Открыт/Сделка/Закрыт)', 'Вид активности', 'Статус доступен для выбора (Да/Нет)'];

var TURN_HEADERS = [
  'ИНН', 'Оборот конкурентов, мес (PS)', 'Банки конкурентов', 'Доля СБП', 'Доля e-com',
  'Оборот Сбер, мес', 'Комиссионный доход Сбер', 'Файл PS', 'Файл Сбер', 'Обновлено'
];
var COMP_HEADERS = ['ИНН', 'Краткое название', 'Полное название', 'ОГРН', 'Регион', 'Статус', 'Источник', 'Обновлено'];
var LOG_HEADERS = ['Время', 'Пользователь', 'Действие', 'ID', 'Детали'];
var PLAN_HEADERS = ['Месяц', 'ФИО', 'План сделок, шт', 'План объёма сделок, ₽', 'Изменено', 'Кто изменил'];

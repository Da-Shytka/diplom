/**
 * Человекочитаемые метки для статусов проверок.
 */
export function statusLabel(status) {
  const map = {
    ok:                         'Всё верно',
    ok_collective:              'Коллективное — ОК',
    ok_no_greeting:             'Обращение отсутствует',
    ok_no_fio:                  'Организация без физлица',
    skip:                       'Пропущено',
    error_gender:               'Неверный род',
    error_mismatch:             'Инициалы не совпадают',
    error_personal_to_multiple: 'Личное при нескольких адресатах',
    error_exception:            'Ошибка обработки',
    llm_error:                  'Ошибка LLM',
  }
  return map[status] || status
}

/**
 * ok* → 'ok' | error* → 'error' | иначе → 'warn'
 */
export function statusKind(status) {
  if (!status) return 'warn'
  if (status.startsWith('ok'))    return 'ok'
  if (status.startsWith('error')) return 'error'
  return 'warn'
}
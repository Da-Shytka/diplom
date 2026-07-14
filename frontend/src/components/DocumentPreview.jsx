import { Tooltip, Tag, Typography, Space } from 'antd'
import { WarningOutlined } from '@ant-design/icons'

const { Text } = Typography

/**
 * Отображает письмо в виде бумажного документа.
 * Если highlight_greeting === true — текст обращения подсвечивается
 * красным, при наведении показывается тултип с перечнем ошибок.
 */
export function DocumentPreview({ result }) {
  const {
    file,
    addressee,
    fio = [],
    greeting,
    highlight_greeting,
    check_gender,
    check_initials,
  } = result

  // Собираем список ошибок для тултипа
  const errorMessages = []
  if (check_gender?.status?.startsWith('error')) {
    errorMessages.push(check_gender.details)
  }
  if (check_initials?.status?.startsWith('error')) {
    errorMessages.push(check_initials.details)
  }

  const tooltipContent = errorMessages.length > 0 ? (
    <div style={{ maxWidth: 320 }}>
      {errorMessages.map((msg, i) => (
        <div key={i} style={{ marginBottom: i < errorMessages.length - 1 ? 6 : 0 }}>
          <WarningOutlined style={{ marginRight: 6, color: '#ff7875' }} />
          {msg}
        </div>
      ))}
    </div>
  ) : null

  // Рендер текста обращения
  const greetingNode = greeting ? (
    highlight_greeting ? (
      <Tooltip
        title={tooltipContent}
        color="#1f1f1f"
        overlayStyle={{ maxWidth: 360 }}
        placement="topLeft"
      >
        <span className="greeting-highlight">
          {greeting}
        </span>
      </Tooltip>
    ) : (
      <span className="greeting-ok">{greeting}</span>
    )
  ) : (
    <Text type="secondary" italic>Обращение не найдено</Text>
  )

  return (
    <div className="letter-paper animate-fade-up">

      {/* Шапка */}
      <div className="letter-header">
        <div className="letter-sender">
          <div style={{ fontSize: 11, color: '#aaa', textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: 2 }}>
            Файл
          </div>
          <div>{file}</div>
        </div>

        <div className="letter-addressee">
          {addressee ? (
            addressee.split('\n').map((line, i) => {
              const trimmed = line.trim()
              if (!trimmed) return null
              const isFio = fio.some(f => trimmed.includes(f.surname))
              return (
                <p key={i} className={isFio ? 'fio-line' : ''}>
                  {trimmed}
                </p>
              )
            })
          ) : (
            <Text type="secondary" italic style={{ fontSize: 13 }}>Адресат не найден</Text>
          )}
        </div>
      </div>

      <hr className="letter-divider" />

      {/* Обращение */}
      <div className="letter-greeting-wrap">
        {greetingNode}
      </div>

      {/* Заглушка тела письма */}
      <div className="letter-body-placeholder">
        <p />
        <p />
        <p />
        <p />
      </div>

      {/* ФИО */}
      {fio.length > 0 && (
        <div style={{ marginTop: 20, paddingTop: 16, borderTop: '1px dashed #eee' }}>
          <div style={{ fontSize: 11, color: '#aaa', textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: 8 }}>
            Найденные ФИО
          </div>
          <Space wrap size={[6, 6]}>
            {fio.map((f, i) => (
              <Tag key={i} color="blue" style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 12 }}>
                {f.raw}
              </Tag>
            ))}
          </Space>
        </div>
      )}
    </div>
  )
}

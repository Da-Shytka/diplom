import { FileTextOutlined, BankOutlined, CheckOutlined } from '@ant-design/icons'

const modes = [
  {
    value: 'outgoing_letter', title: 'Исходящее письмо', Icon: FileTextOutlined,
  },
  {
    value: 'court_letter', title: 'Судебный документ', Icon: BankOutlined,
  },
]

export function CheckModePicker({ value, onChange }) {
  return (
    <section className="mode-section" aria-labelledby="mode-heading">
      <div className="mode-heading">
        <span className="step-label">01 / РЕЖИМ ПРОВЕРКИ</span>
        <h2 id="mode-heading">Какой документ проверяем?</h2>
      </div>
      <fieldset className="mode-grid">
        <legend className="sr-only">Режим проверки документа</legend>
        {modes.map(({ value: mode, title, Icon }) => (
          <label key={mode} className={`mode-card ${value === mode ? 'is-selected' : ''}`}>
            <input type="radio" name="document-mode" aria-label={title} value={mode} checked={value === mode} onChange={() => onChange(mode)} />
            <span className="mode-icon" aria-hidden="true"><Icon /></span>
            <span className="mode-title">{title}</span>
            <span className="mode-radio" aria-hidden="true">{value === mode && <CheckOutlined />}</span>
          </label>
        ))}
      </fieldset>
      <div className="upload-heading"><span className="step-label">02 / ДОКУМЕНТ</span><h3>Добавьте файл для проверки</h3></div>
    </section>
  )
}

import { useState } from 'react'
import { Button, Input, Space, Typography, message } from 'antd'
import { FormOutlined, ClearOutlined } from '@ant-design/icons'

const { TextArea } = Input
const { Text } = Typography

export function TextCheckPanel({ onCheck }) {
  const [text, setText] = useState('')

  function submit() {
    if (!text.trim()) {
      message.warning('Вставьте текст для проверки')
      return
    }
    onCheck(text)
  }

  return (
    <div className="text-check-panel">
      <TextArea
        value={text}
        onChange={(event) => setText(event.target.value)}
        placeholder="Вставьте текст письма или отдельный фрагмент для проверки грамотности, пунктуации, стиля и нормативных ссылок"
        autoSize={{ minRows: 11, maxRows: 18 }}
        style={{
          fontFamily: "'Times New Roman', Georgia, serif",
          fontSize: 15,
          lineHeight: 1.7,
          resize: 'vertical',
        }}
      />

      <Space style={{ marginTop: 14, width: '100%', justifyContent: 'space-between' }}>
        <Text type="secondary" style={{ fontSize: 12 }}>
          Можно проверять текст без шапки и подписи.
        </Text>
        <Space>
          <Button icon={<ClearOutlined />} onClick={() => setText('')} disabled={!text}>
            Очистить
          </Button>
          <Button type="primary" icon={<FormOutlined />} onClick={submit}>
            Проверить текст
          </Button>
        </Space>
      </Space>
    </div>
  )
}

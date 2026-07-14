import { useState } from 'react'
import {
  Layout, Typography, Button, Badge, Space, Spin, Result, Tabs,
} from 'antd'
import {
  FileTextOutlined, ReloadOutlined,
} from '@ant-design/icons'

import { useDocumentCheck } from './hooks/useDocumentCheck'
import { DropZone } from './components/DropZone'
import { FullDocumentPreview } from './components/FullDocumentPreview'
import { TextCheckPanel } from './components/TextCheckPanel'

const { Header, Sider, Content } = Layout
const { Title, Text } = Typography

/* ─── Хелперы ──────────────────────────────────────────────────── */

function statusBadge(status) {
  if (!status) return <Badge status="default" text="—" />
  if (status.startsWith('ok'))    return <Badge status="success" text={statusLabel(status)} />
  if (status.startsWith('error')) return <Badge status="error"   text={statusLabel(status)} />
  return <Badge status="warning" text={statusLabel(status)} />
}

function statusLabel(s) {
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
  return map[s] || s
}

/* ─── App ──────────────────────────────────────────────────────── */

export default function App() {
  const { result, loading, error, checkDocument, checkText, reset } = useDocumentCheck()
  const [fileName, setFileName] = useState(null)
  const [mode, setMode] = useState('file')

  function handleFile(file) {
    setFileName(file.name)
    setMode('file')
    checkDocument(file)
  }

  function handleText(text) {
    setFileName('вставленный текст')
    setMode('text')
    checkText(text)
  }

  function resetAll() {
    reset()
    setFileName(null)
  }

  return (
    <Layout style={{ minHeight: '100vh', fontFamily: "'Geologica', system-ui, sans-serif" }}>

      {/* ─── Сайдбар ─────────────────────────────────────────── */}
      <Sider
        width={220}
        theme="light"
        style={{
          borderRight: '1px solid #f0f0f0',
          padding: '24px 0',
        }}
      >
        {/* Логотип */}
        <div style={{ padding: '0 20px 28px', display: 'flex', alignItems: 'center', gap: 10 }}>
          <div style={{
            width: 36, height: 36, background: '#e6f4ff', borderRadius: 8,
            display: 'flex', alignItems: 'center', justifyContent: 'center',
          }}>
            <FileTextOutlined style={{ color: '#1677ff', fontSize: 18 }} />
          </div>
          <span style={{ fontWeight: 700, fontSize: 15, letterSpacing: '-0.01em' }}>
            LetterCheck
          </span>
        </div>

        {/* Подсказка */}
        <div style={{ padding: '0px 20px 0', marginTop: 'auto' }}>
          <Text type="secondary" style={{ fontSize: 12, lineHeight: 1.55 }}>
            Загрузите .docx файл или вставьте обычный текст для проверки.
          </Text>
        </div>
      </Sider>

      {/* ─── Основной контент ─────────────────────────────────── */}
      <Layout>

        {/* Топбар */}
        <Header style={{
          background: '#fff',
          borderBottom: '1px solid #f0f0f0',
          padding: '0 32px',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          height: 56,
        }}>
          <div>
            <Title level={5} style={{ margin: 0, fontWeight: 700 }}>
              Проверка письма:
              <Text type="secondary" style={{ fontSize: 12, fontFamily: "'JetBrains Mono', monospace" }}>
                &nbsp;{fileName}
              </Text>
            </Title>
          </div>

          {(result || error) && (
            <Button
              icon={<ReloadOutlined />}
              onClick={resetAll}
            >
              Новая проверка
            </Button>
          )}
        </Header>

        <Content style={{ padding: '28px 32px', overflow: 'auto' }}>

          {/* Загрузка */}
          {!result && !loading && !error && (
            <Tabs
              activeKey={mode}
              onChange={setMode}
              items={[
                {
                  key: 'file',
                  label: 'Файл .docx',
                  children: <DropZone onFile={handleFile} />,
                },
                {
                  key: 'text',
                  label: 'Текст',
                  children: <TextCheckPanel onCheck={handleText} />,
                },
              ]}
            />
          )}

          {/* Спиннер */}
          {loading && (
            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 16, padding: '80px 0' }}>
              <Spin size="large" />
              <Text type="secondary">Анализируем документ…</Text>
            </div>
          )}

          {/* Ошибка сети/сервера */}
          {error && (
            <Result
              status="error"
              title="Ошибка обработки"
              subTitle={error}
              extra={
                <Button onClick={resetAll}>
                  Попробовать снова
                </Button>
              }
            />
          )}

          {/* Результат */}
          {result && (
            <div className="animate-fade-up">
              <Space direction="vertical" size={20} style={{ width: '100%' }}>

                <FullDocumentPreview result={result} />

              </Space>
            </div>
          )}

        </Content>
      </Layout>
    </Layout>
  )
}

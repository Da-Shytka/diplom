// frontend/src/components/FullDocumentPreview.jsx
import React from 'react'
import { Card, Tooltip, Typography, Tag, Alert } from 'antd'
import {
  CheckCircleOutlined,
  CloseCircleOutlined,
  InfoCircleOutlined,
  WarningOutlined,
} from '@ant-design/icons'

const { Text } = Typography

const POPOVER_TEXT_LIMIT = 140
const POPOVER_CORRECTION_LIMIT = 90
const POPOVER_ITEMS_LIMIT = 3

function compactPopoverText(value, limit = POPOVER_TEXT_LIMIT) {
  if (!value) return ''
  const normalized = String(value).replace(/\s+/g, ' ').trim()
  if (normalized.length <= limit) return normalized
  return `${normalized.slice(0, limit).trim()}...`
}

// ── Вспомогательная функция: вставляем подсветку в произвольный текст ──────
function HighlightedText({ text, highlights, baseColor, borderColor, icon }) {
  if (!text) return null
  if (!highlights || highlights.length === 0) return <span>{text}</span>

  // Сортируем по позиции, убираем пересечения
  const sorted = [...highlights]
    .sort((a, b) => a.start - b.start)
    .filter((h, i, arr) => i === 0 || h.start >= arr[i - 1].end)

  const parts = []
  let cursor = 0

  for (const h of sorted) {
    if (h.start > cursor) {
      parts.push(<span key={`t-${cursor}`}>{text.slice(cursor, h.start)}</span>)
    }
    const fragment = text.slice(h.start, h.end)
    parts.push(
      <Tooltip
        key={`h-${h.start}`}
        title={
          <div style={{ maxWidth: 340 }}>
            <div style={{ marginBottom: 6, fontWeight: 'bold' }}>
              {h.title || 'Ошибка'}
            </div>

            {h.items?.length > 0 ? (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                {h.items.slice(0, POPOVER_ITEMS_LIMIT).map((item, index) => (
                  <div key={index}>
                    {item.title && (
                      <div style={{ fontSize: 12, fontWeight: 600 }}>
                        {compactPopoverText(item.title, 70)}
                      </div>
                    )}
                    {item.description && (
                      <div style={{ fontSize: 12 }}>
                        {compactPopoverText(item.description)}
                      </div>
                    )}
                    {item.corrected && (
                      <div style={{ marginTop: 4, fontSize: 12 }}>
                        <span style={{ opacity: 0.7 }}>Исправление: </span>
                        <span style={{ fontWeight: 'bold', color: '#52c41a' }}>
                          {compactPopoverText(item.corrected, POPOVER_CORRECTION_LIMIT)}
                        </span>
                      </div>
                    )}
                  </div>
                ))}
                {h.items.length > POPOVER_ITEMS_LIMIT && (
                  <div style={{ fontSize: 12, opacity: 0.75 }}>
                    Еще {h.items.length - POPOVER_ITEMS_LIMIT} замеч. ниже
                  </div>
                )}
              </div>
            ) : (
              <>
                {h.description && (
                  <div style={{ fontSize: 12 }}>
                    {compactPopoverText(h.description)}
                  </div>
                )}
                {h.corrected && (
                  <div style={{ marginTop: 6, fontSize: 12 }}>
                    <span style={{ opacity: 0.7 }}>Исправление: </span>
                    <span style={{ fontWeight: 'bold', color: '#52c41a' }}>
                      {compactPopoverText(h.corrected, POPOVER_CORRECTION_LIMIT)}
                    </span>
                  </div>
                )}
              </>
            )}
          </div>
        }
        placement="top"
      >
        <mark
          style={{
            background: baseColor || '#ffccc7',
            borderBottom: `2px solid ${borderColor || '#ff4d4f'}`,
            cursor: 'help',
            padding: '1px 3px',
            borderRadius: 3,
            fontWeight: 500,
          }}
        >
          {fragment}
          {icon && <WarningOutlined style={{ marginLeft: 4, fontSize: 11, color: borderColor || '#ff4d4f' }} />}
        </mark>
      </Tooltip>
    )
    cursor = h.end
  }

  if (cursor < text.length) {
    parts.push(<span key="tail">{text.slice(cursor)}</span>)
  }

  return <>{parts}</>
}

// ── Шапка документа (таблица) ───────────────────────────────────────────────
function DocumentHeader({ headerRows }) {
  if (!headerRows || headerRows.length === 0) return null

  return (
    <div className="document-header">
      {headerRows.map((cells, rowIdx) => (
        <div
          key={rowIdx}
          style={{
            display: 'grid',
            gridTemplateColumns: cells.length === 1 ? '1fr' : '1fr 1fr',
            gap: 18,
            marginBottom: 5,
            alignItems: 'start',
          }}
        >
          {cells.map((cell, cellIdx) => (
            <div
              key={cellIdx}
              style={{
                fontSize: 14,
                lineHeight: 1.5,
                // Правая колонка (адресат) — по правому краю
                textAlign: cells.length > 1 && cellIdx === cells.length - 1 ? 'right' : 'left',
                whiteSpace: 'pre-line',
                overflowWrap: 'anywhere',
              }}
            >
              {cell}
            </div>
          ))}
        </div>
      ))}
    </div>
  )
}

// ── Обращение ───────────────────────────────────────────────────────────────
function DocumentGreeting({ greeting, hasError, checkGender, checkInitials }) {
  if (!greeting) return null

  const highlights = []

  if (hasError) {
    const greetingErrors = []

    if (checkGender?.status?.startsWith('error')) {
      greetingErrors.push({
        title: 'Ошибка в роде обращения',
        description: checkGender.details,
        corrected: checkGender.corrected,
      })
    }

    if (checkInitials?.status?.startsWith('error')) {
      greetingErrors.push(...(checkInitials.errors || [{
        title: 'Ошибка в обращении',
        description: checkInitials.details,
        corrected: checkInitials.corrected,
      }]))
    }

    highlights.push({
      start: 0,
      end: greeting.length,
      title: 'Ошибка в обращении',
      items: greetingErrors,
    })
  }

  return (
    <p
      style={{
        marginBottom: 8,
        lineHeight: 1.8,
        textAlign: 'justify',
      }}
    >
      <HighlightedText
        text={greeting}
        highlights={highlights}
        baseColor="#fffbe6"
        borderColor="#faad14"
        icon
      />
    </p>
  )
}

// ── Основной текст ──────────────────────────────────────────────────────────
function DocumentBody({ paragraphs, spellingErrors }) {
  if (!paragraphs || paragraphs.length === 0) return null

  // Переводим ошибки в highlights для каждого параграфа
  const getParaHighlights = (paraText, paraOffset) => {
    if (!spellingErrors || spellingErrors.length === 0) return []
    return spellingErrors
      .filter(e => e.start >= paraOffset && e.start < paraOffset + paraText.length)
      .map(e => ({
        start: e.start - paraOffset,
        end: Math.min(e.end - paraOffset, paraText.length),
        title: errorLabel(e.type),
        description: e.description,
        corrected: e.corrected,
      }))
  }

  let offset = 0
  let textParagraphsSeen = 0
  return (
    <div>
      {paragraphs.map((para, idx) => {
        if (para.empty) {
          return <div key={idx} style={{ height: 12 }} />
        }
        if (textParagraphsSeen > 0) offset += 1
        const paraOffset = offset
        offset += para.text.length
        textParagraphsSeen += 1

        const highlights = getParaHighlights(para.text, paraOffset)
        const style = para.style || {}

        return (
          <p
            key={idx}
            style={{
              marginBottom: 8,
              lineHeight: 1.8,
              textAlign: style.alignment || 'justify',
              paddingLeft: style.indent_left ? `${style.indent_left}pt` : undefined,
              fontWeight: style.bold ? 600 : 400,
              textIndent: style.first_line_indent ? `${style.first_line_indent}pt` : '1.5em',
            }}
          >
            <HighlightedText
              text={para.text}
              highlights={highlights}
              baseColor="#fffbe6"
              borderColor="#faad14"
              icon
            />
          </p>
        )
      })}
    </div>
  )
}

// ── Подпись ─────────────────────────────────────────────────────────────────
function DocumentSignature({ paragraphs }) {
  if (!paragraphs || paragraphs.length === 0) return null
  return (
    <div className="document-signature">
      {paragraphs.map((para, idx) => (
        <div
          key={idx}
          style={{
            fontSize: 14,
            lineHeight: 1.6,
            textAlign: para.style?.alignment || 'left',
          }}
        >
          {para.text}
        </div>
      ))}
    </div>
  )
}

function DocumentExecutor({ paragraphs }) {
  if (!paragraphs || paragraphs.length === 0) return null
  return (
    <div className="document-executor">
      {paragraphs.map((para, idx) => (
        <div key={idx}>
          {para.text}
        </div>
      ))}
    </div>
  )
}

function errorLabel(type) {
  const map = {
    spelling: 'Орфография',
    punctuation: 'Пунктуация',
    style: 'Стиль',
    normative: 'Нормативная ссылка',
  }
  return map[type] || 'Текст'
}

function requiredStatusColor(status) {
  const map = {
    ok: '#52c41a',
    missing: '#ff4d4f',
    optional_present: '#1677ff',
    optional_absent: '#bfbfbf',
  }
  return map[status] || '#bfbfbf'
}

function requiredStatusLabel(block) {
  if (block.status === 'ok') return 'найден'
  if (block.status === 'missing') return 'не найден'
  if (block.status === 'optional_present') return 'найден'
  return 'не требуется'
}

function RequiredBlocksPanel({ checkRequiredBlocks }) {
  if (!checkRequiredBlocks || checkRequiredBlocks.status === 'skip') return null

  const rows = checkRequiredBlocks.blocks || []
  const classification = checkRequiredBlocks.classification || {}
  const hasErrors = checkRequiredBlocks.status === 'has_errors'

  return (
    <div className="required-blocks-panel">
      <div className="required-blocks-head">
        <div>
          <div className="required-blocks-title">Обязательные блоки</div>
          <div className="required-blocks-subtitle">
            Тип документа: <b>{checkRequiredBlocks.document_type_label || checkRequiredBlocks.document_type}</b>
            {classification.confidence && (
              <Tag color={classification.confidence === 'high' ? 'green' : 'gold'} style={{ marginLeft: 8 }}>
                {classification.confidence === 'high' ? 'уверенно' : 'по признакам'}
              </Tag>
            )}
          </div>
        </div>
        <Tag color={hasErrors ? 'error' : 'success'} style={{ margin: 0 }}>
          {hasErrors ? 'Есть пропуски' : 'Все на месте'}
        </Tag>
      </div>

      {classification.reasons?.length > 0 && (
        <div className="required-blocks-reasons">
          <InfoCircleOutlined />
          <span>{classification.reasons.join('; ')}</span>
        </div>
      )}

      {hasErrors && (
        <Alert
          type="error"
          showIcon
          message={checkRequiredBlocks.details}
          style={{ marginBottom: 12 }}
        />
      )}

      <div className="required-blocks-grid">
        {rows.map((block) => {
          const color = requiredStatusColor(block.status)
          const requiredText = block.required ? 'обязателен' : 'дополнительно'
          return (
            <div
              key={block.key}
              className={`required-block-row ${block.status === 'missing' ? 'is-missing' : ''}`}
            >
              <div className="required-block-icon" style={{ color }}>
                {block.present ? <CheckCircleOutlined /> : <CloseCircleOutlined />}
              </div>
              <div className="required-block-main">
                <div className="required-block-name">
                  {block.label}
                  <Tag color={block.required ? 'blue' : 'default'} style={{ marginLeft: 8 }}>
                    {requiredText}
                  </Tag>
                </div>
                <div className="required-block-state" style={{ color }}>
                  {requiredStatusLabel(block)}
                </div>
                {block.evidence && (
                  <div className="required-block-evidence">
                    {block.evidence}
                  </div>
                )}
              </div>
            </div>
          )
        })}
      </div>
    </div>
  )
}

// ── Легенда ошибок ───────────────────────────────────────────────────────────
function ErrorLegend({ checkGender, checkInitials, checkSpelling, checkRequiredBlocks }) {
  const genderErr = checkGender?.status?.startsWith('error')
  const initialsErr = checkInitials?.status?.startsWith('error')
  const spellingErr = checkSpelling?.status === 'has_errors'
  const requiredErr = checkRequiredBlocks?.status === 'has_errors'

  if (!genderErr && !initialsErr && !spellingErr && !requiredErr) return null

  return (
    <div style={{
      marginTop: 16,
      padding: '12px 16px',
      background: '#fffbe6',
      border: '1px solid #ffe58f',
      borderRadius: 8,
      display: 'flex',
      flexDirection: 'column',
      gap: 6,
    }}>
      {genderErr && (
        <div style={{ display: 'flex', alignItems: 'flex-start', gap: 8 }}>
          <Tag color="error" style={{ flexShrink: 0, marginTop: 1 }}>Род</Tag>
          <Text style={{ fontSize: 13 }}>{checkGender.details}</Text>
        </div>
      )}
      {initialsErr && (
        <div style={{ display: 'flex', alignItems: 'flex-start', gap: 8 }}>
          <Tag color="error" style={{ flexShrink: 0, marginTop: 1 }}>Инициалы</Tag>
          <Text style={{ fontSize: 13 }}>{checkInitials.details}</Text>
        </div>
      )}
      {spellingErr && (
        <>
          <div style={{ display: 'flex', alignItems: 'flex-start', gap: 8 }}>
            <Tag color="warning" style={{ flexShrink: 0, marginTop: 1 }}>Текст</Tag>
            <Text style={{ fontSize: 13 }}>{checkSpelling.details}</Text>
          </div>
          {checkSpelling.errors?.slice(0, 6).map((error, index) => (
            <div key={index} style={{ display: 'flex', alignItems: 'flex-start', gap: 8 }}>
              <Tag color={error.type === 'normative' ? 'blue' : 'gold'} style={{ flexShrink: 0, marginTop: 1 }}>
                {errorLabel(error.type)}
              </Tag>
              <Text style={{ fontSize: 13 }}>{error.description}</Text>
            </div>
          ))}
        </>
      )}
      {requiredErr && (
        <div style={{ display: 'flex', alignItems: 'flex-start', gap: 8 }}>
          <Tag color="error" style={{ flexShrink: 0, marginTop: 1 }}>Блоки</Tag>
          <Text style={{ fontSize: 13 }}>{checkRequiredBlocks.details}</Text>
        </div>
      )}
    </div>
  )
}

// ── Основной компонент ───────────────────────────────────────────────────────
export function FullDocumentPreview({ result }) {
  const {
    blocks,
    check_gender,
    check_initials,
    check_spelling,
    check_required_blocks,
    mode,
  } = result

  if (!blocks) {
    return <Text type="secondary">Нет данных для отображения</Text>
  }

  const spellingErrors = check_spelling?.errors || []

  const hasGreetingError =
    check_gender?.status?.startsWith('error') ||
    check_initials?.status?.startsWith('error')

  return (
    <Card bodyStyle={{ padding: 0 }}>
      {/* Документ */}
      <div
        className="document-paper"
        style={{
          background: 'white',
          padding: '40px 48px',
          borderRadius: 8,
          fontFamily: "'Times New Roman', Georgia, serif",
          fontSize: 15,
          lineHeight: 1.7,
          color: '#1a1a1a',
          maxHeight: 640,
          overflowY: 'auto',
        }}
      >
        {mode !== 'text' && <DocumentHeader headerRows={blocks.header_rows} />}
        <DocumentGreeting
          greeting={blocks.greeting}
          hasError={hasGreetingError}
          checkGender={check_gender}
          checkInitials={check_initials}
        />
        <DocumentBody
          paragraphs={blocks.body_paragraphs}
          spellingErrors={spellingErrors}
        />
        <DocumentSignature paragraphs={blocks.signature_paragraphs} />
        <DocumentExecutor paragraphs={blocks.executor_paragraphs} />
      </div>

      {/* Легенда */}
      <div style={{ padding: '0 16px 16px' }}>
        {mode !== 'text' && (
          <RequiredBlocksPanel checkRequiredBlocks={check_required_blocks} />
        )}
        <ErrorLegend
          checkGender={check_gender}
          checkInitials={check_initials}
          checkSpelling={check_spelling}
          checkRequiredBlocks={check_required_blocks}
        />
      </div>
    </Card>
  )
}

import { useState, useCallback } from 'react'

/**
 * Хук для отправки .docx файла на бэкенд и получения результата проверки.
 */
export function useDocumentCheck() {
  const [result, setResult]   = useState(null)
  const [loading, setLoading] = useState(false)
  const [error, setError]     = useState(null)

  const checkDocument = useCallback(async (file) => {
    setLoading(true)
    setError(null)
    setResult(null)

    const formData = new FormData()
    formData.append('file', file)

    try {
      const resp = await fetch('/check', {
        method: 'POST',
        body: formData,
      })

      if (!resp.ok) {
        const data = await resp.json().catch(() => ({}))
        throw new Error(data.detail || `Ошибка сервера: ${resp.status}`)
      }

      const data = await resp.json()
      setResult(data)
    } catch (err) {
      setError(err.message || 'Неизвестная ошибка')
    } finally {
      setLoading(false)
    }
  }, [])

  const checkText = useCallback(async (text) => {
    setLoading(true)
    setError(null)
    setResult(null)

    try {
      const resp = await fetch('/check-text', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ text }),
      })

      if (!resp.ok) {
        const data = await resp.json().catch(() => ({}))
        throw new Error(data.detail || `Ошибка сервера: ${resp.status}`)
      }

      const data = await resp.json()
      setResult(data)
    } catch (err) {
      setError(err.message || 'Неизвестная ошибка')
    } finally {
      setLoading(false)
    }
  }, [])

  const reset = useCallback(() => {
    setResult(null)
    setError(null)
    setLoading(false)
  }, [])

  return { result, loading, error, checkDocument, checkText, reset }
}

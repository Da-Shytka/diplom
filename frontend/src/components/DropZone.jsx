// frontend/src/components/DropZone.jsx
import { Upload, message } from 'antd'
import { InboxOutlined } from '@ant-design/icons'

const { Dragger } = Upload

export function DropZone({ onFile }) {
  const uploadProps = {
    name: 'file',
    multiple: false,
    accept: '.docx',
    showUploadList: false,
    beforeUpload: (file) => {
      // Проверка расширения файла
      const isDocx = file.name.endsWith('.docx')
      if (!isDocx) {
        message.error('Можно загружать только файлы формата .docx')
        return false
      }
      
      // Проверка размера (максимум 10MB)
      const isLessThan10MB = file.size / 1024 / 1024 < 10
      if (!isLessThan10MB) {
        message.error('Файл должен быть меньше 10MB')
        return false
      }
      
      // Вызываем callback с файлом
      onFile(file)
      return false // Предотвращаем автоматическую загрузку
    },
    onChange: (info) => {
      if (info.file.status === 'error') {
        message.error('Ошибка при загрузке файла')
      }
    },
  }

  return (
    <Dragger {...uploadProps}>
      <p className="ant-upload-drag-icon">
        <InboxOutlined />
      </p>
      <p className="ant-upload-text">Нажмите или перетащите файл сюда</p>
      <p className="ant-upload-hint">
        Поддерживаются только файлы .docx. Максимальный размер: 10MB.
      </p>
    </Dragger>
  )
}
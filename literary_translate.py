#!/usr/bin/env python3
"""
Литературный переводчик документов
Основной исполняемый файл для перевода .docx документов с английского на русский
"""

import sys
import os
import argparse
import asyncio
from pathlib import Path
from typing import Optional

# Добавляем текущую директорию в путь Python
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import config
from document_processor import DocumentProcessor
from text_chunker import TextChunker
from translator import DocumentTranslator
from logger_config import setup_logging, TranslationProgress, TranslationLogger


def parse_arguments():
    """Парсит аргументы командной строки"""
    parser = argparse.ArgumentParser(
        description="Литературный переводчик .docx документов с английского на русский",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры использования:
  python literary_translate.py input.docx
  python literary_translate.py input.docx --xml
  python literary_translate.py input.docx --chunk-size 1500
  python literary_translate.py input.docx --log-level DEBUG
        """
    )
    
    parser.add_argument(
        'input_file',
        nargs='?',
        help='Путь к входному .docx файлу'
    )
    
    parser.add_argument(
        '--xml',
        action='store_true',
        help='Дополнительно сохранить переведенный текст в XML формате'
    )
    
    parser.add_argument(
        '--chunk-size',
        type=int,
        default=config.chunk_size,
        help=f'Максимальный размер блока текста для перевода (по умолчанию: {config.chunk_size})'
    )
    
    parser.add_argument(
        '--log-level',
        choices=['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'],
        default=config.log_level,
        help=f'Уровень логирования (по умолчанию: {config.log_level})'
    )
    
    parser.add_argument(
        '--test-api',
        action='store_true',
        help='Протестировать подключение к API и выйти'
    )
    
    return parser.parse_args()


def generate_output_filename(input_file: str) -> str:
    """Генерирует имя выходного файла на основе входного"""
    input_path = Path(input_file)
    # Создаем имя: translated_original_name.docx
    output_name = f"translated_{input_path.stem}.docx"
    output_path = input_path.parent / output_name
    return str(output_path)


def validate_input_file(input_file: str) -> bool:
    """Проверяет корректность входного файла"""
    input_path = Path(input_file)
    
    if not input_path.exists():
        print(f"❌ Входной файл не найден: {input_file}")
        return False
    
    if not input_path.is_file():
        print(f"❌ Указанный путь не является файлом: {input_file}")
        return False
    
    if input_path.suffix.lower() != '.docx':
        print(f"❌ Поддерживаются только .docx файлы. Получен: {input_path.suffix}")
        return False
    
    return True


def validate_output_file(output_file: str) -> bool:
    """Проверяет корректность выходного файла"""
    output_path = Path(output_file)
    
    # Проверяем, что директория существует или может быть создана
    output_dir = output_path.parent
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        print(f"❌ Не удалось создать директорию для выходного файла: {e}")
        return False
    
    return True


def test_api_connection() -> bool:
    """Тестирует подключение к API"""
    try:
        translator = DocumentTranslator()
        if translator.api_translator.test_connection():
            print("✅ Подключение к OpenRouter API успешно!")
            return True
        else:
            print("❌ Не удалось подключиться к OpenRouter API")
            return False
    except Exception as e:
        print(f"❌ Ошибка при тестировании API: {e}")
        return False


async def run_translation(args):
    """Асинхронная функция для перевода документа"""
    # Создаем объект для логирования
    logger = TranslationLogger()
    
    # Генерируем имя выходного файла
    output_file = generate_output_filename(args.input_file)
    
    # Логируем начало работы
    logger.log_start(args.input_file, output_file)
    
    try:
        # Создаем процессор документов
        doc_processor = DocumentProcessor()
        
        # Загружаем документ
        if not doc_processor.load_document(args.input_file):
            logger.log_error("Не удалось загрузить документ")
            sys.exit(1)
        
        # --- НОВАЯ АСИНХРОННАЯ ЛОГИКА ---
        new_document = await doc_processor.process_and_translate_async()

        if not new_document:
            logger.log_error("Не удалось обработать и перевести документ")
            sys.exit(1)
            
        if not doc_processor.save_document_with_images(new_document, output_file):
            logger.log_error("Не удалось сохранить переведенный документ")
            sys.exit(1)
        
        doc_processor.cleanup_temp_files()
        
        # Сохраняем XML файл если нужно
        xml_file = None
        if args.xml:
            xml_file = str(Path(output_file).with_suffix('.xml'))
            if not doc_processor.save_as_xml(xml_file):
                logger.log_error("Не удалось сохранить XML файл")
                xml_file = None
        
        # Логируем успешное завершение
        logger.log_success(output_file, xml_file)
        
    except KeyboardInterrupt:
        logger.log_error("Операция была прервана пользователем")
        sys.exit(1)
    except Exception as e:
        logger.log_error(f"Неожиданная ошибка: {str(e)}")
        sys.exit(1)


def main():
    """Основная функция программы"""
    # Парсим аргументы
    args = parse_arguments()
    
    # Настраиваем логирование
    setup_logging(args.log_level)
    
    # Если нужно только протестировать API
    if args.test_api:
        sys.exit(0 if test_api_connection() else 1)
    
    # Проверяем, что указан входной файл
    if not args.input_file:
        print("❌ Не указан входной файл")
        sys.exit(1)
    
    # Проверяем входные данные
    if not validate_input_file(args.input_file):
        sys.exit(1)
    
    # Генерируем имя выходного файла
    output_file = generate_output_filename(args.input_file)
    
    if not validate_output_file(output_file):
        sys.exit(1)
    
    try:
        # Запускаем асинхронную функцию перевода
        asyncio.run(run_translation(args))
        
    except KeyboardInterrupt:
        print("❌ Операция была прервана пользователем")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Неожиданная ошибка: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main() 
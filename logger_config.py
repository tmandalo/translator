"""
Конфигурация логирования для переводчика
"""

import logging
import sys
from typing import Optional, Callable
from pathlib import Path
from datetime import datetime

from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TimeRemainingColumn, TaskProgressColumn
from rich.logging import RichHandler
from rich.text import Text
from rich.panel import Panel
from rich.table import Table

from config import config


class ColoredFormatter(logging.Formatter):
    """Форматтер для цветного логирования"""
    
    COLORS = {
        'DEBUG': 'cyan',
        'INFO': 'green',
        'WARNING': 'yellow',
        'ERROR': 'red',
        'CRITICAL': 'bold red'
    }
    
    def format(self, record):
        log_color = self.COLORS.get(record.levelname, 'white')
        record.levelname = f"[{log_color}]{record.levelname}[/{log_color}]"
        return super().format(record)


def setup_logging(log_level: str = None) -> logging.Logger:
    """
    Настраивает логирование для приложения
    
    Args:
        log_level: Уровень логирования (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        
    Returns:
        Настроенный logger
    """
    if log_level is None:
        log_level = config.log_level
    
    # Создаем консоль Rich
    console = Console()
    
    # Создаем директорию для логов если она не существует
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)
    
    # Создаем имя файла лога с датой
    log_file = log_dir / f"translation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    
    # Настраиваем root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(getattr(logging, log_level.upper()))
    
    # Очищаем существующие handlers
    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)
    
    # Создаем Rich handler для консоли
    rich_handler = RichHandler(
        console=console,
        rich_tracebacks=True,
        show_time=True,
        show_level=True,
        show_path=False
    )
    rich_handler.setLevel(getattr(logging, log_level.upper()))
    
    # Создаем file handler
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    
    # Создаем форматтеры
    console_formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%H:%M:%S'
    )
    
    file_formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Применяем форматтеры
    rich_handler.setFormatter(console_formatter)
    file_handler.setFormatter(file_formatter)
    
    # Добавляем handlers
    root_logger.addHandler(rich_handler)
    root_logger.addHandler(file_handler)
    
    return root_logger


class TranslationProgress:
    """Класс для отображения прогресса перевода"""
    
    def __init__(self, total_chunks: int):
        self.console = Console()
        self.total_chunks = total_chunks
        self.progress = None
        self.task_id = None
        self.successful = 0
        self.failed = 0
        self.start_time = None
        
    def __enter__(self):
        self.progress = Progress(
            SpinnerColumn(),
            TextColumn("[bold blue]Переводим документ..."),
            BarColumn(),
            TaskProgressColumn(),
            TextColumn("•"),
            TextColumn("[green]Успешно: {task.fields[successful]}"),
            TextColumn("[red]Ошибок: {task.fields[failed]}"),
            TextColumn("•"),
            TimeRemainingColumn(),
            console=self.console,
            transient=False
        )
        
        self.progress.start()
        self.task_id = self.progress.add_task(
            "Перевод...",
            total=self.total_chunks,
            successful=0,
            failed=0
        )
        
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.progress:
            self.progress.stop()
    
    def update(self, completed: int, total: int, success: bool):
        """Обновляет прогресс"""
        if success:
            self.successful += 1
        else:
            self.failed += 1
        
        if self.progress and self.task_id is not None:
            self.progress.update(
                self.task_id,
                completed=completed,
                successful=self.successful,
                failed=self.failed
            )
    
    def show_summary(self, stats: dict):
        """Показывает итоговую статистику"""
        # Создаем таблицу с результатами
        table = Table(title="📊 Результаты перевода")
        table.add_column("Метрика", style="cyan")
        table.add_column("Значение", style="green")
        
        # Добавляем строки
        table.add_row("Всего блоков", str(stats.get('total_chunks', 0)))
        table.add_row("Переведено успешно", str(stats.get('successful_chunks', 0)))
        table.add_row("Ошибок", str(stats.get('failed_chunks', 0)))
        table.add_row("Процент успеха", f"{stats.get('success_rate', 0)*100:.1f}%")
        table.add_row("Время выполнения", f"{stats.get('total_processing_time', 0):.1f}с")
        table.add_row("Использовано токенов", str(stats.get('total_tokens_used', 0)))
        
        # Создаем панель
        panel = Panel(table, title="✅ Перевод завершен", border_style="green")
        
        self.console.print()
        self.console.print(panel)


class TranslationLogger:
    """Класс для логирования процесса перевода"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.console = Console()
    
    def log_start(self, input_file: str, output_file: str):
        """Логирует начало перевода"""
        self.console.print()
        self.console.print(Panel(
            f"🔄 Начинаем перевод документа\n"
            f"📄 Входной файл: {input_file}\n"
            f"📄 Выходной файл: {output_file}",
            title="🚀 Литературный переводчик",
            border_style="blue"
        ))
        
        self.logger.info(f"Начинаем перевод: {input_file} -> {output_file}")
    
    def log_document_stats(self, stats: dict):
        """Логирует статистику документа"""
        message_parts = [
            f"📊 Статистика документа:",
            f"• Всего элементов: {stats.get('total_elements', 0)}",
            f"• Параграфов: {stats.get('paragraphs', 0)}",
            f"• Таблиц: {stats.get('tables', 0)}",
            f"• Изображений: {stats.get('images', 0)}",
            f"• Символов: {stats.get('total_characters', 0)}",
            f"• Средний размер элемента: {stats.get('average_element_size', 0):.0f}"
        ]
        
        # Добавляем детали об изображениях если есть
        if stats.get('images', 0) > 0:
            message_parts.append("\n🖼️ Детали изображений:")
            
            if 'formats' in stats:
                formats_info = ", ".join([f"{fmt}: {count}" for fmt, count in stats['formats'].items()])
                message_parts.append(f"• Форматы: {formats_info}")
            
            if 'inline_images' in stats:
                message_parts.append(f"• Встроенных: {stats.get('inline_images', 0)}")
                
            if 'floating_images' in stats:
                message_parts.append(f"• Плавающих: {stats.get('floating_images', 0)}")
        
        self.console.print()
        self.console.print(Panel(
            "\n".join(message_parts),
            title="📋 Анализ документа",
            border_style="cyan"
        ))
        
        self.logger.info(f"Статистика документа: {stats}")
    
    def log_chunk_stats(self, stats: dict):
        """Логирует статистику разбивки на блоки"""
        self.console.print()
        self.console.print(Panel(
            f"🔨 Разбивка на блоки:\n"
            f"• Всего блоков: {stats.get('total_chunks', 0)}\n"
            f"• Средний размер блока: {stats.get('average_chunk_size', 0):.0f} символов\n"
            f"• Максимальный блок: {stats.get('max_chunk_size', 0)} символов\n"
            f"• Минимальный блок: {stats.get('min_chunk_size', 0)} символов\n"
            f"• Полных параграфов: {stats.get('complete_paragraphs', 0)}",
            title="🧩 Подготовка к переводу",
            border_style="yellow"
        ))
        
        self.logger.info(f"Статистика блоков: {stats}")
    
    def log_error(self, error: str):
        """Логирует ошибку"""
        self.console.print()
        self.console.print(Panel(
            f"❌ Ошибка: {error}",
            title="💥 Ошибка",
            border_style="red"
        ))
        
        self.logger.error(error)
    
    def log_success(self, output_file: str, xml_file: str = None):
        """Логирует успешное завершение"""
        message = f"✅ Перевод завершен успешно!\n📄 Сохранен файл: {output_file}"
        
        if xml_file:
            message += f"\n📄 XML файл: {xml_file}"
        
        self.console.print()
        self.console.print(Panel(
            message,
            title="🎉 Успех",
            border_style="green"
        ))
        
        self.logger.info(f"Перевод завершен: {output_file}") 